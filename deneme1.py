import os
import glob
import re
import json
import time
import requests
import pandas as pd
import pdfplumber
import extract_msg
import email 
from email import policy
from email.parser import BytesParser
import csv
from datetime import datetime
from bs4 import BeautifulSoup
import dateparser
import glob
import difflib
import re
import warnings
import os
import streamlit as st
import tempfile

st.set_page_config(page_title="Tarife İşleme", layout="wide")


def final_pol_correction():
    """
    Programın en sonunda çalışır. 
    Tüm çıktı CSV dosyalarındaki 'POL' sütununu kontrol eder ve 
    belirtilen liman isimlerini tmaxx standartlarına çevirir.
    """
    print("\n--- ⚓ Final POL Düzenleme İşlemi Başlıyor ---")
    
    # Mevcut dizindeki tüm CSV dosyalarını bul
    all_csvs = glob.glob("*.csv")
    
    # Değişim kuralları (Büyük/Küçük harf duyarsız kontrol için)
    corrections = {
        "CHILUNG": "KEELUNG",
        "NAGOYA": "Nagoya, Aichi",
        "KUPANG": "Kupang, Timor"
    }

    for csv_file in all_csvs:
        try:
            # Dosyayı oku (Dosya yapısına uygun olarak ayırıcı ; ve utf-8-sig)
            df = pd.read_csv(csv_file, sep=';', encoding='utf-8-sig', engine='python')
            
            if 'POL' not in df.columns:
                continue

            changed = False
            
            # POL sütunundaki her satırı kontrol et
            def apply_rules(val):
                nonlocal changed
                if pd.isna(val): return val
                
                original_val = str(val).strip()
                upper_val = original_val.upper()
                
                # Kuralları uygula
                new_val = original_val
                if upper_val == "CHILUNG":
                    new_val = "KEELUNG"
                elif upper_val == "NAGOYA":
                    new_val = "Nagoya, Aichi"
                elif upper_val == "KUPANG":
                    new_val = "Kupang, Timor"
                
                if new_val != original_val:
                    changed = True
                    return new_val
                return original_val

            df['POL'] = df['POL'].apply(apply_rules)

            if changed:
                df.to_csv(csv_file, index=False, sep=';', encoding='utf-8-sig')
                print(f"   ✅ {csv_file} dosyası güncellendi (POL düzeltmeleri yapıldı).")
                
        except Exception as e:
            print(f"   ⚠️ {csv_file} işlenirken hata oluştu: {e}")

    print("--- 🏁 Final POL Düzenleme Tamamlandı ---\n")


def clean_text(val):
    if pd.isna(val): return ""
    t = str(val)
    t = t.replace('\ufeff', '').replace('\u200b', '').replace('\xa0', ' ')
    return t.strip(' "\'')

def apply_tmaxx_corrections_to_all_csvs(tmaxx_excel_path):
    print("\n--- 🔍 Tmaxx Liman İsimleri (Akıllı Eşleştirme) Başlıyor ---")
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    
    try:
        if tmaxx_excel_path.endswith('.csv'):
            df_tmaxx = pd.read_csv(tmaxx_excel_path, sep=None, engine='python')
        else:
            df_tmaxx = pd.read_excel(tmaxx_excel_path)
            
        df_tmaxx.columns = df_tmaxx.columns.str.strip()
        
        tmaxx_map = {}
        for _, row in df_tmaxx.iterrows():
            if 'Ad' not in row or pd.isna(row['Ad']): continue
            
            orijinal_isim = clean_text(row['Ad'])
            if not orijinal_isim: continue
            
            key = orijinal_isim.upper()
            
            # 1. Normal Eşleşme
            if key not in tmaxx_map:
                tmaxx_map[key] = orijinal_isim
            elif tmaxx_map[key].isupper() and not orijinal_isim.isupper():
                tmaxx_map[key] = orijinal_isim
                
            # 2. Boşluksuz Eşleşme
            bosluksuz = key.replace(" ", "")
            if bosluksuz not in tmaxx_map or (tmaxx_map[bosluksuz].isupper() and not orijinal_isim.isupper()):
                tmaxx_map[bosluksuz] = orijinal_isim
                
            # 3. Uluslararası Kod
            if 'Uluslararası Kod' in row and pd.notna(row['Uluslararası Kod']):
                tmaxx_map[clean_text(row['Uluslararası Kod']).upper()] = orijinal_isim
                
        # 4. YENİ: EĞİK ÇİZGİ (/) VE PARANTEZ () YAKALAYICI (Alias Sistemi)
        eklenecek_aliaslar = {}
        for key, orijinal_isim in tmaxx_map.items():
            # / ile ayrılmış isimler (Örn: JIUJIANG /SHAHEJIE)
            if '/' in key:
                for part in key.split('/'):
                    p = part.strip()
                    if len(p) >= 3 and p not in tmaxx_map and p not in eklenecek_aliaslar:
                        eklenecek_aliaslar[p] = orijinal_isim
                        
            # Parantez içindeki isimler (Örn: SHUNDE (RONGQI))
            if '(' in key and ')' in key:
                ic_kisim = re.search(r'\((.*?)\)', key)
                if ic_kisim:
                    p = ic_kisim.group(1).strip()
                    if len(p) >= 3 and p not in tmaxx_map and p not in eklenecek_aliaslar:
                        eklenecek_aliaslar[p] = orijinal_isim
                        
        # Yan isimleri ana hafızaya kat
        tmaxx_map.update(eklenecek_aliaslar)
            
        print(f"   ✅ Referans listeden {len(set(tmaxx_map.values()))} adet ana liman ve yan isim hafızaya alındı.")
            
    except Exception as e:
        print(f"   ⚠️ Tmaxx referans dosyası okunamadı: {e}")
        return

    valid_names_upper = sorted([k for k in tmaxx_map.keys() if isinstance(k, str) and len(k)>3], key=len, reverse=True)
    all_csvs = glob.glob("*.csv")
    
    baz_dosya_adi = os.path.basename(tmaxx_excel_path)
    if baz_dosya_adi in all_csvs:
        all_csvs.remove(baz_dosya_adi)

    eslesme_hafizasi = {}
    
    def akilli_eslestir(val):
        temizlenmis_orijinal = clean_text(val)
        if not temizlenmis_orijinal: return val
        
        val_str = temizlenmis_orijinal.upper()
        
        if val_str in eslesme_hafizasi:
            return eslesme_hafizasi[val_str]
            
        sonuc = val 
        bosluksuz_val = val_str.replace(" ", "")
        
        if val_str in tmaxx_map: 
            sonuc = tmaxx_map[val_str]
        elif bosluksuz_val in tmaxx_map:
            sonuc = tmaxx_map[bosluksuz_val]
        else:
            temiz_val = re.sub(r'\([^)]*\)', '', val_str).strip()
            temiz_val = temiz_val.split(',')[0].strip()
            temiz_val = temiz_val.split('/')[0].strip()
            
            if temiz_val in tmaxx_map: 
                sonuc = tmaxx_map[temiz_val]
            elif temiz_val.replace(" ", "") in tmaxx_map:
                sonuc = tmaxx_map[temiz_val.replace(" ", "")]
            else:
                kelimeler = temiz_val.split()
                if len(kelimeler) > 0:
                    yari_uzunluk = len(kelimeler) // 2
                    if yari_uzunluk > 0 and kelimeler[:yari_uzunluk] == kelimeler[yari_uzunluk:]:
                        tekil_metin = " ".join(kelimeler[:yari_uzunluk])
                        if tekil_metin in tmaxx_map:
                            sonuc = tmaxx_map[tekil_metin]
                
                if sonuc == val:
                    for valid_isim in valid_names_upper:
                        if len(valid_isim) >= 4 and len(temiz_val) >= 4:
                            if re.search(r'\b' + re.escape(valid_isim) + r'\b', temiz_val):
                                sonuc = tmaxx_map[valid_isim]
                                break 
                            
                            if valid_isim.startswith(temiz_val + " ") or valid_isim.startswith(temiz_val + "-"):
                                sonuc = tmaxx_map[valid_isim]
                                break
                                
                            if valid_isim == temiz_val + " PORT" or valid_isim == temiz_val + " PT":
                                sonuc = tmaxx_map[valid_isim]
                                break

                if sonuc == val:
                    ad_listesi = list(set(str(v).upper() for v in tmaxx_map.values()))
                    benzerler = difflib.get_close_matches(temiz_val, ad_listesi, n=1, cutoff=0.88)
                    if benzerler: 
                        sonuc = tmaxx_map.get(benzerler[0], val)
                        
        eslesme_hafizasi[val_str] = sonuc
        return sonuc

    for csv_file in all_csvs:
        try:
            df = pd.read_csv(csv_file, sep=None, engine='python')
            changed = False 
            
            hedef_kelimeler = ['POL', 'POD', 'PORT OF LOADING', 'PORT OF DISCHARGE']
            hedef_sutunlar = []
            for col in df.columns:
                temiz_col = clean_text(col).upper()
                if temiz_col in hedef_kelimeler:
                    hedef_sutunlar.append(col)
            
            if not hedef_sutunlar:
                continue
            
            temiz_sutun_isimleri = [clean_text(c) for c in hedef_sutunlar]
            print(f"   📂 İşleniyor ({csv_file}) -> Bulunan Hedef Sütunlar: {temiz_sutun_isimleri}")
            
            degisiklik_ozeti = set()
            
            for col in hedef_sutunlar:
                for idx, row in df.iterrows():
                    eski_deger = df.at[idx, col]
                    yeni_deger = akilli_eslestir(eski_deger)
                    
                    if pd.notna(eski_deger) and str(eski_deger).strip() != str(yeni_deger).strip():
                        df.at[idx, col] = yeni_deger
                        changed = True
                        degisiklik_ozeti.add(f"'{eski_deger}' ➡️ '{yeni_deger}'")
                        
            for degisiklik in degisiklik_ozeti:
                print(f"      ✏️ Değişti: {degisiklik}")
                        
            if changed:
                df.to_csv(csv_file, index=False, sep=';', encoding='utf-8-sig')
                
        except Exception as e:
            print(f"   ⚠️ Hata ({csv_file}): {e}")
            
    print("--- 🏁 Tmaxx Eşleştirme İşlemi Tamamlandı ---\n")

# ============================================================
# [YENİ] LİMAN KODU EŞLEŞTİRME LİSTESİ (VIA İÇİN)
# ============================================================
PORT_CODE_MAP = {
    # -- SENİN EKLEDİKLERİN --
    "CNDLC": "DALIAN",
    "CNTAO": "QINGDAO",
    "CNNSA": "NANSHA",
    "VNCMP": "CAI MEP",      # Vietnam
    "CNSHK": "SHEKOU",
    "CNXMN": "XIAMEN",
    "CNYTN": "YANTIAN",
    "HKHKG": "HONG KONG",
    "TWKHH": "KAOHSIUNG",    # Tayvan
    "CNNGB": "NINGBO",
    "KRPUS": "BUSAN", "PUSAN": "BUSAN", "BUSAN": "BUSAN",
    "CNSNH": "SHANGHAI", "SHANGHAI": "SHANGHAI",
    "CNNGB": "NINGBO", "CNSGB": "NINGBO", "NINGBO": "NINGBO",
    "CNQIN": "QINGDAO", "QINGDAO": "QINGDAO",
    "CNSAD": "DA CHAN BAY", "DA CHAN BAY": "DA CHAN BAY",
    "VNTCT": "CAI MEP", "VNHCM": "CAI MEP", "CAI MEP": "CAI MEP",
    "THLEM": "LAEM CHABANG",
    "CNTXG": "XINGANG", "CNYTN": "YANTIAN",
    "HKHKG": "HONG KONG", "SGSIN": "SINGAPORE",
    
    # -- ESKİ LİSTEDEKİLER --
    "SGSIN": "SINGAPORE",
    "EGPSD": "PORT SAID",
    "GRPIR": "PIRAEUS",
    "MYPKG": "PORT KLANG",
    "ESALG": "ALGECIRAS",
    "DEHAM": "HAMBURG",
    "NLRTM": "ROTTERDAM",
    "BEANT": "ANTWERP",
    "CNSHA": "SHANGHAI",
    "CNNING": "NINGBO",      # Alternatif kod
    "KRPUS": "BUSAN",
    "ITSPE": "LA SPEZIA",
    "ITGOA": "GENOA",
    "MTMAR": "MARSAXLOKK",
    "SAJED": "JEDDAH",
    "AEJEA": "JEBEL ALI",
    "LKCMB": "COLOMBO"
}

# ============================================================
# [YENİ] COSCO İÇİN ETS MUAFİYET KONTROLÜ
# ============================================================
def is_china_region_port(pol_name):
    """
    Verilen POL (Çıkış Limanı) Çin, Tayvan, Hong Kong veya Macao'da mı?
    Eğer öyleyse True döner (ETS Eklenmemeli).
    """
    if not pol_name: return False
    pol = pol_name.upper().strip()

    # 1. Ülke Kodları (CN: Çin, HK: Hong Kong, TW: Tayvan, MO: Macao)
    if pol.startswith(("CN", "HK", "TW", "MO")):
        return True

    # 2. İsim Bazlı Kontrol (Kod yerine şehir ismi yazıyorsa)
    china_regions = [
        "DALIAN", "QINGDAO", "TIANJIN", "QINHUANGDAO", "SHANGHAI", "NINGBO",
        "LIANYUNGANG", "NANJING", "NANTONG", "ZHANGJIAGANG", "ZHENJIANG", "YANGZHOU",
        "CHANGSHU", "TAICANG", "CHANGZHOU", "CHONGQING", "LUZHOU", "CHANGSHA", "YICHANG",
        "YUEYANG", "ANQING", "WUHAN", "JIUJIANG", "JIANGXI", "WUHU", "WENZHOU", "XIAMEN",
        "FUZHOU", "FUJIAN", "NANSHA", "SANSHUI NEW PORT", "FOSHAN", "NANHAI", "HONG KONG",
        "SHEKOU", "YANTIAN", "BEIJIAO", "GAOMING", "HUANGPU", "JIANGMEN", "SHUNDE", "QINZHOU",
        "RONGQI", "SHENWAN", "WUZHOU", "XIAOLAN", "XINHUI", "ZHANJIANG", "ZHONGSHAN", "ZHUHAI",
        "SHANTOU", "KAOHSIUNG", "KEELUNG", "TAIPEI", "TAICHUNG", "TAOYUAN"
    ]
    
    for city in china_regions:
        if city in pol:
            return True
            
    return False

import glob
import re
import csv
import pandas as pd
import extract_msg
from datetime import datetime
from email.parser import BytesParser
from email import policy

def cma_special_logic_processor():
    """
    Klasörde CMA maili ve TAO/TAD exceli varsa devreye girer.
    GH ve SUEZ rotalarını tespit edip iki ayrı dosya üretir.
    """
    # 1. Dosya Tespiti
    all_files = glob.glob("*.*")
    cma_mail_file = None
    cma_excel_file = None

    for f in all_files:
        if f.lower().endswith(('.msg', '.eml')) and 'cma' in f.lower():
            cma_mail_file = f
            break
    
    for f in all_files:
        if f.lower().endswith(('.xlsx', '.csv')) and ('tao' in f.lower() or 'tad' in f.lower()):
            cma_excel_file = f
            break

    if not cma_mail_file or not cma_excel_file:
        return

    print("\n" + "="*60)
    print(f"🚀 CMA ÖZEL MODU DEVREYE GİRDİ")
    print(f"   Mail Dosyası : {cma_mail_file}")
    print(f"   Tarife Dosyası: {cma_excel_file}")
    print("="*60 + "\n")

    try:
        # --- SABİTLER ---
        FE_MP_CODES = {
            'CNDLC': 'DALIAN', 'CNNGB': 'NINGBO', 'CNNSA': 'NANSHA', 'CNSHA': 'SHANGHAI',
            'CNSHK': 'SHEKOU', 'CNTAO': 'QINGDAO', 'CNTXG': 'TIANJIN', 'CNXMN': 'XIAMEN',
            'CNYTN': 'YANTIAN', 'HKHKG': 'HONG KONG', 'TWKHH': 'KAOHSIUNG',
            'KRPUS': 'BUSAN', 'SGSIN': 'SINGAPORE', 'MYPKG': 'PORT KELANG',
            'MYTPP': 'TANJUNG PELEPAS', 'VNVUT': 'VUNG TAU', 'IDJKT': 'JAKARTA',
            'IDSUB': 'SURABAYA', 'VNSGN': 'HO CHI MINH'
        }
        
        TR_MAIN_PORTS = ['AMBARLI', 'MERSIN', 'ALIAGA', 'ISKENDERUN', 'GEMLIK']
        ENS_FEE = 27.0

        # --- ADIM 1: MAİLDEN BAZ FİYATLARI ÇEKME ---
        base_rates = {}
        email_body = ""
        
        if cma_mail_file.lower().endswith('.msg'):
            msg_obj = extract_msg.Message(cma_mail_file)
            email_body = msg_obj.body if msg_obj.body else ""
            if (not email_body or len(email_body.strip()) < 10) and msg_obj.htmlBody:
                email_body = msg_obj.htmlBody.decode('iso-8859-9', errors='ignore')
            msg_obj.close()
        else:
            with open(cma_mail_file, 'rb') as f:
                msg_obj = BytesParser(policy=policy.default).parse(f)
                body_part = msg_obj.get_body(preferencelist=('plain', 'html'))
                email_body = body_part.get_content() if body_part else ""

        if email_body is None:
            email_body = ""

        # Gelişmiş Regex: GH ve SUEZ fiyatlarını yakalar
        gh_match = re.search(r'via\s+GH[\s\S]*?\$(\d+)\s*/\s*\$?(\d+)', email_body, re.IGNORECASE)
        if gh_match:
            base_rates['GH'] = {'20': float(gh_match.group(1)), '40': float(gh_match.group(2))}
        
        suez_match = re.search(r'via\s+SUEZ[\s\S]*?\$(\d+)\s*/\s*\$?(\d+)', email_body, re.IGNORECASE)
        if suez_match:
            base_rates['SUEZ'] = {'20': float(suez_match.group(1)), '40': float(suez_match.group(2))}

        if not base_rates:
            print(f"⚠️ UYARI: {cma_mail_file} içinde 'via GH' veya 'via SUEZ' fiyatları yakalanamadı.")
            return
        else:
            print(f"   ✅ Baz Fiyatlar Bulundu: {base_rates}")

        # --- ADIM 2: EXCEL / CSV OKUMA ---
        df = None
        header_row_idx = None
        
        if cma_excel_file.lower().endswith('.csv'):
            df_temp = pd.read_csv(cma_excel_file, header=None, nrows=20, sep=None, engine='python')
        else:
            df_temp = pd.read_excel(cma_excel_file, header=None, nrows=20)
        
        for i, row in df_temp.iterrows():
            row_str = str(row.values).upper()
            if "POL" in row_str and "20'ST" in row_str:
                header_row_idx = i
                break
        
        if header_row_idx is None: header_row_idx = 6

        if cma_excel_file.lower().endswith('.csv'):
            df = pd.read_csv(cma_excel_file, skiprows=header_row_idx, sep=None, engine='python')
        else:
            df = pd.read_excel(cma_excel_file, skiprows=header_row_idx)

        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if 'POL' not in df.columns:
            print(f"❌ HATA: 'POL' sütunu bulunamadı.")
            return

        # Referans çözümleme haritası
        data_map = {}
        for idx, row in df.iterrows():
            if pd.isna(row.get('POL')): continue
            pol_key = str(row['POL']).strip().upper()
            code_key = str(row.get('CODE', '')).strip().upper()
            row_data = {
                'POL': pol_key,
                'VIA': str(row.get('PTS', '')),
                '20': row.get("20'ST"),
                '40': row.get("40'ST"),
                '40HC': row.get("40'HC")
            }
            data_map[pol_key] = row_data
            if code_key and code_key != 'NAN': data_map[code_key] = row_data

        def resolve_via_name(code):
            if pd.isna(code) or str(code).lower() in ['nan', 'direct', '-']:
                return ""
            parts = str(code).split('/')
            names = [FE_MP_CODES.get(p.strip().upper(), p.strip().upper()) for p in parts]
            return "/".join(names)

        def parse_price(val):
            if pd.isna(val): return None
            s = str(val).strip().upper()
            if s == 'BP': return 0.0
            if s in ['N/A', 'NO SOLUTION', '-']: return None
            if 'REFER' in s or 'SEE' in s: return f"REF:{s}"
            try: return float(s)
            except: return None

        # --- ADIM 3: VERİ İŞLEME VE AYIRMA ---
        # Her rota için ayrı liste tutuyoruz
        separated_results = { 'GH': [], 'SUEZ': [] }

        for idx, row in df.iterrows():
            pol_raw = row.get('POL')
            if pd.isna(pol_raw): continue
            
            pol_name = str(pol_raw).strip().upper()
            pts_val = row.get('PTS')
            p20 = parse_price(row.get("20'ST"))
            p40 = parse_price(row.get("40'ST"))
            p40hc = parse_price(row.get("40'HC"))
            
            # Referans Çözümleme
            if isinstance(p20, str) and p20.startswith("REF:"):
                match_ref = re.search(r'(?:REFER TO|SEE)\s+([A-Z0-9/\s,]+)', p20)
                if match_ref:
                    potential_targets = re.split(r'[/\s,]+', match_ref.group(1).strip())
                    for t in potential_targets:
                        t_clean = t.strip()
                        if t_clean in data_map:
                            target_row = data_map[t_clean]
                            p20 = parse_price(target_row['20'])
                            p40 = parse_price(target_row['40'])
                            p40hc = parse_price(target_row['40HC'])
                            pts_val = target_row['VIA']
                            break
            
            if not isinstance(p20, (int, float)) or not isinstance(p40, (int, float)):
                continue

            via_final = resolve_via_name(pts_val)
            via_list = [v.strip() for v in via_final.split('/') if v.strip()] if via_final else [""]

            for route_type, base_vals in base_rates.items():
                total_20 = base_vals['20'] + p20 + ENS_FEE
                total_40 = base_vals['40'] + p40 + ENS_FEE
                current_p40hc = p40hc if isinstance(p40hc, (int, float)) else p40
                total_40hc = base_vals['40'] + current_p40hc + ENS_FEE

                # POL yanına parantez eklemiyoruz, sadece liman ismi
                pol_display = pol_name

                for tr_pod in TR_MAIN_PORTS:
                    for v_port in via_list:
                        separated_results[route_type].append({
                            'POL': pol_display,
                            'VIA': v_port,
                            'POD': tr_pod,
                            'CURR': "USD",
                            'FREETIME': "16",
                            'TYPE_1': "20 DC", 'AMOUNT_1': total_20,
                            'TYPE_2': "40 DC", 'AMOUNT_2': total_40,
                            'TYPE_3': "40 HC", 'AMOUNT_3': total_40hc
                        })

        # --- ADIM 4: KAYDET (AYRI DOSYALAR) ---
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        FINAL_HEADER = ["POL", "VIA", "POD", "TOCITY", "CURR", "FREETIME", "TRANSIT", 
                        "CUSTOMERDESCRIPTION", "TYPE", "AMOUNT", "TYPE", "AMOUNT", "TYPE", "AMOUNT"]
        
        cols_map = ["POL", "VIA", "POD", "TOCITY", "CURR", "FREETIME", "TRANSIT", "CUSTOMERDESCRIPTION",
                    "TYPE_1", "AMOUNT_1", "TYPE_2", "AMOUNT_2", "TYPE_3", "AMOUNT_3"]

        found_any = False
        for route_name, results in separated_results.items():
            if not results:
                continue
            
            found_any = True
            out_name = f"CMA_{route_name}_OUTPUT_{timestamp}.csv"
            
            df_out = pd.DataFrame(results)
            # Eksik sütunları tamamla
            for col in cols_map:
                if col not in df_out.columns: df_out[col] = ""
            
            df_final = df_out[cols_map]
            
            with open(out_name, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(FINAL_HEADER)
                for _, row in df_final.iterrows():
                    writer.writerow(row.tolist())

            print(f"✅ İŞLEM TAMAMLANDI. {route_name} dosyası: {out_name} ({len(df_final)} satır)")

        if not found_any:
            print("⚠️ İşlenecek veri bulunamadı.")

    except Exception as e:
        import traceback
        print(f"❌ CMA Modülü Hatası: {e}")
        print(traceback.format_exc())





# =============================================================================
# HAPAG-LLOYD ÖZEL ENTEGRASYON MODÜLÜ
# =============================================================================


def load_hapag_addons():
    """Hapag Add-On dosyasını yükler ve sözlük formatına çevirir."""
    addon_file = "Hapag Add-On.xlsx"
    addons = {}
    
    if os.path.exists(addon_file):
        try:
            # Excel'i oku (İlk sayfa varsayılıyor)
            df_addon = pd.read_excel(addon_file)
            # Kolon isimlerini temizle
            df_addon.columns = [c.strip().lower() for c in df_addon.columns]
            
            # Origin City ve fiyatları al (Örnek kolon adlarına göre)
            # Excel'deki kolonlar: Origin City, 20', 40' varsayıyoruz
            for _, row in df_addon.iterrows():
                city = str(row.get('origin city', '')).strip().upper()
                price_20 = row.get("20'", 0)
                price_40 = row.get("40'", 0)
                
                # Sayısal olmayan değerleri 0 yap
                try: price_20 = float(price_20)
                except: price_20 = 0.0
                try: price_40 = float(price_40)
                except: price_40 = 0.0
                
                if city:
                    addons[city] = {'20': price_20, '40': price_40}
            print(f"   ✅ Add-On listesi yüklendi: {len(addons)} liman.")
        except Exception as e:
            print(f"   ⚠️ Add-On dosyası okunurken hata: {e}")
    else:
        print("   ⚠️ Hapag Add-On.xlsx bulunamadı, ek masraflar eklenmeyecek.")
    
    return addons

def process_hapag_special(file_path):
    """
    Hapag EML/MSG dosyalarını işler.
    
    YENİ GÜNCELLEME (FINAL):
    1. 'Unique Date Count': Bir tabloda 2'den fazla farklı tarih varsa (örn: 04.02, 08.02, 14.02)
       bu tabloyu 'Wrapper' (Dış Çerçeve) olarak işaretler ve İŞLEMEZ.
    2. 'Superset Cleaning': Veri işlendikten sonra, eğer bir tarih aralığı (04-14)
       başka bir tarih aralığını (04-08) kapsıyorsa ve fiyatlar aynıysa,
       geniş olan (hatalı) aralığı veritabanından siler.
    """
    import os
    import pandas as pd
    import re
    from bs4 import BeautifulSoup
    import csv as import_csv
    from email import policy
    from email.parser import BytesParser
    import extract_msg
    from datetime import datetime

    print(f"   ⚓ HAPAG İşleniyor (Final - Çakışma Korumalı): {os.path.basename(file_path)}")
    
    # 1. ADD-ON YÜKLEME
    addon_file = "Hapag Add-On.xlsx"
    addon_map = {} 
    if os.path.exists(addon_file):
        try:
            df_addon = pd.read_excel(addon_file)
            df_addon.columns = [str(c).strip().upper() for c in df_addon.columns]
            col_origin = next((c for c in df_addon.columns if "ORIGIN" in c or "CITY" in c), None)
            col_20 = next((c for c in df_addon.columns if "20" in c), None)
            col_40 = next((c for c in df_addon.columns if "40" in c), None)
            if col_origin:
                for _, row in df_addon.iterrows():
                    pol = str(row[col_origin]).strip().upper()
                    if not pol or pol == "NAN": continue
                    try: p20 = float(row[col_20]) if col_20 and pd.notna(row[col_20]) else 0
                    except: p20 = 0
                    try: p40 = float(row[col_40]) if col_40 and pd.notna(row[col_40]) else 0
                    except: p40 = 0
                    addon_map[pol] = {'20': p20, '40': p40}
        except: pass

    # 2. MAİL İÇERİĞİNİ OKU
    html_content = ""
    try:
        if file_path.lower().endswith(".eml"):
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            body = msg.get_body(preferencelist=('html'))
            if body: html_content = body.get_content()
            else:
                body = msg.get_body(preferencelist=('plain'))
                if body: html_content = body.get_content()
        elif file_path.lower().endswith(".msg"):
            msg = extract_msg.Message(file_path)
            html_content = msg.htmlBody
            msg.close()
    except Exception as e:
        print(f"      ❌ Okuma Hatası: {e}")
        return

    if not html_content: return

    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    target_pods = ["ISTANBUL", "IZMIT", "ALIAGA", "MERSIN"]

    all_collected_rows = []

    # 3. VERİ TOPLAMA
    for tbl in tables:
        # KORUMA 1: Nested Table
        if tbl.find("table"):
            continue

        # Metin Temizliği (Header vs.)
        rows_text = []
        for tr in tbl.find_all('tr'):
            row_str = tr.get_text(" ", strip=True).upper()
            if any(x in row_str for x in ["DATE:", "SENT:", "RECEIVED:", "SUBJECT:", "FROM:", "TO:"]):
                continue
            rows_text.append(row_str)
        
        tbl_text_clean = " ".join(rows_text)

        if "QFP 2" not in tbl_text_clean and "QFP2" not in tbl_text_clean:
            continue
        
        if not any(pod in tbl_text_clean for pod in target_pods):
            continue

        # Tarihleri Çek
        dates_in_table = re.findall(r'(\d{2}\.\d{2}\.\d{4})', tbl_text_clean)
        
        if not dates_in_table:
            continue

        # --- KORUMA 2: UNIQUE TARİH SAYISI KONTROLÜ (ÇOK ÖNEMLİ) ---
        # Bir tabloda en fazla 2 farklı tarih olabilir (Start ve End).
        # Eğer 3 veya daha fazla farklı tarih varsa (04.02, 08.02, 09.02...), 
        # bu tablo birden fazla dönemi kapsayan bir 'Hayalet Tablo'dur. ATLA.
        unique_dates = sorted(list(set(dates_in_table)))
        if len(unique_dates) > 2:
            # print(f"      🛡️ Dış tablo atlandı (Farklı tarih sayısı: {len(unique_dates)})")
            continue
        # -----------------------------------------------------------

        validity_str = dates_in_table[0]
        start_date_obj = None
        end_date_obj = None

        try:
            if len(dates_in_table) >= 2:
                # Daima ilk iki tarihi al (Zaten unique kontrolü 2'den fazla olamaz diyor)
                start_d = datetime.strptime(dates_in_table[0], "%d.%m.%Y")
                end_d = datetime.strptime(dates_in_table[1], "%d.%m.%Y")
                
                day_diff = (end_d - start_d).days
                
                if day_diff < 3:
                    continue
                
                validity_str = f"{dates_in_table[0]}-{dates_in_table[1]}"
                start_date_obj = start_d
                end_date_obj = end_d
            else:
                validity_str = f"Valid-{dates_in_table[0]}"
                start_date_obj = datetime.strptime(dates_in_table[0], "%d.%m.%Y")
                end_date_obj = start_date_obj # Tek tarih varsa bitiş=başlangıç varsayalım
        except:
            continue

        # Fiyatları Çek
        rows = tbl.find_all('tr')
        qfp_prices = {}
        found_qfp = False
        
        for tr in rows:
            cells = tr.find_all(['td', 'th'])
            row_text = " ".join([c.get_text(strip=True).upper() for c in cells])
            
            if "QFP 2" in row_text or "QFP2" in row_text:
                all_numbers = []
                for cell in cells:
                    txt = cell.get_text(strip=True)
                    val_clean = re.sub(r"[^\d]", "", txt)
                    if val_clean:
                        val_int = int(val_clean)
                        if val_int > 100: all_numbers.append(val_int)
                
                if len(all_numbers) >= 8:
                    chunk_size = len(all_numbers) // 4
                    current_idx = 0
                    for pod in target_pods:
                        if current_idx + 1 < len(all_numbers):
                            p20 = all_numbers[current_idx]
                            p40 = all_numbers[current_idx + 1]
                            qfp_prices[pod] = {'20': p20, '40': p40}
                            current_idx += chunk_size
                    found_qfp = True
                    break 
        
        if found_qfp:
            for origin, addons in addon_map.items():
                for pod, bases in qfp_prices.items():
                    try:
                        total_20 = int(bases['20'] + addons['20'])
                        total_40 = int(bases['40'] + addons['40'])
                        
                        all_collected_rows.append({
                            'VALIDITY': validity_str,
                            'START_DATE': start_date_obj,
                            'END_DATE': end_date_obj,
                            'POL': origin,
                            'VIA': "",
                            'POD': pod,
                            'TOCITY': "",
                            'CURR': "USD",
                            'FREETIME': "16",
                            'TRANSIT': "",
                            'DESC': "",
                            'TYPE1': "20 DC",
                            'PRICE1': total_20,
                            'TYPE2': "40 DC",
                            'PRICE2': total_40,
                            'TYPE3': "40 HC",
                            'PRICE3': total_40
                        })
                    except: continue

    if not all_collected_rows:
        print("      ⚠️ Veri bulunamadı.")
        return

    df = pd.DataFrame(all_collected_rows)
    df.drop_duplicates(inplace=True)

    # --- KORUMA 3: SUPERSET (KAPSAMA) TEMİZLİĞİ ---
    # Eğer "04.02-14.02" gibi bir aralık, "04.02-08.02" gibi daha dar bir aralığı 
    # kapsıyorsa ve fiyatlar aynıysa, geniş olanı (hatalı olanı) sil.
    
    validities = df[['VALIDITY', 'START_DATE', 'END_DATE']].drop_duplicates()
    validities_to_remove = set()

    for _, row_wide in validities.iterrows():
        for _, row_narrow in validities.iterrows():
            if row_wide['VALIDITY'] == row_narrow['VALIDITY']:
                continue
            
            # Tarih Kapsama Kontrolü: Wide kapsar Narrow
            if row_wide['START_DATE'] <= row_narrow['START_DATE'] and \
               row_wide['END_DATE'] >= row_narrow['END_DATE']:
                
                # Sadece tarih yetmez, fiyatların da benzer/aynı olması lazım.
                # Geniş aralıktaki verilerin bir kısmını alıp kontrol edelim.
                wide_data = df[df['VALIDITY'] == row_wide['VALIDITY']]
                narrow_data = df[df['VALIDITY'] == row_narrow['VALIDITY']]
                
                # Eğer aynı limanlar için fiyatlar aynıysa, geniş olan HATALIDIR.
                # Basit hash kontrolü (POL+POD+PRICE)
                is_duplicate = False
                for _, w_row in wide_data.head(3).iterrows():
                    match = narrow_data[
                        (narrow_data['POL'] == w_row['POL']) & 
                        (narrow_data['POD'] == w_row['POD']) & 
                        (narrow_data['PRICE1'] == w_row['PRICE1'])
                    ]
                    if not match.empty:
                        is_duplicate = True
                        break
                
                if is_duplicate:
                    print(f"      🧹 Temizlik: '{row_wide['VALIDITY']}' aralığı, '{row_narrow['VALIDITY']}' aralığının kopyası/kapsayanı olduğu için siliniyor.")
                    validities_to_remove.add(row_wide['VALIDITY'])

    # Hatalı validity'leri ana tablodan çıkar
    if validities_to_remove:
        df = df[~df['VALIDITY'].isin(validities_to_remove)]
    # ----------------------------------------------

    # 4. YAZMA
    for validity, group_df in df.groupby('VALIDITY'):
        csv_output_rows = []
        for _, row in group_df.iterrows():
            csv_output_rows.append([
                row['POL'], row['VIA'], row['POD'], row['TOCITY'], row['CURR'], 
                row['FREETIME'], row['TRANSIT'], row['DESC'],
                row['TYPE1'], row['PRICE1'],
                row['TYPE2'], row['PRICE2'],
                row['TYPE3'], row['PRICE3']
            ])

        fname = f"hapag-{validity}.csv"
        if os.path.exists(fname):
            try: os.remove(fname)
            except: pass
            
        header_str = "POL;VIA;POD;TOCITY;CURR;FREETIME;TRANSIT;CUSTOMERDESCRIPTION;TYPE;AMOUNT;TYPE;AMOUNT;TYPE;AMOUNT"
        try:
            with open(fname, "w", encoding="utf-8-sig", newline="") as f_out:
                writer = import_csv.writer(f_out, delimiter=";")
                f_out.write(header_str + "\n")
                writer.writerows(csv_output_rows)
            print(f"      ✅ Dosya Oluşturuldu: {fname}")
        except: pass



def zim_special_logic_processor(file_path):
    """
    ZIM Mailleri için Gelişmiş İşlemci (V11 - Hybrid Text & Table):
    1. Mailin METİN kısmındaki (POL: ... POD: ...) ana hatları okur (Direct).
    2. HTML TABLOLARDAKİ yan hatları okur (Feeder).
    3. Hepsini birleştirip tek liste döner.
    """

    # PORT_CODE_MAP kontrolü
    if 'PORT_CODE_MAP' not in globals():
        global PORT_CODE_MAP
        PORT_CODE_MAP = {} 

    print(f"\n🚢 ZIM GELİŞMİŞ MODU (V11 - Hybrid): {os.path.basename(file_path)}")

    try:
        msg = extract_msg.Message(file_path)
        body_text = msg.body            # Metin analizi için
        body_html = msg.htmlBody        # Tablo analizi için
        msg.close()
    except Exception as e:
        print(f"❌ ZIM Maili açılamadı: {e}")
        return []

    zim_results = []

    # -------------------------------------------------------------------------
    # ADIM 1: ORTAK VERİLERİ (FİYAT, EKSTRA MASRAF, POD) BULMA
    # -------------------------------------------------------------------------
    
    # A) Ekstra Masraflar (IPS, SMD vb.)
    extra_total = 0.0
    surcharge_keywords = ['IPS', 'SMD', 'ISPS', 'SEAL', 'LSS', 'DOC', 'BAF']
    surcharge_matches = re.findall(r'USD\s*([\d\.]+)\s*/\s*[A-Z]+\s+([A-Z]+)', body_text, re.IGNORECASE)
    
    print("   🔎 [ZIM] Ekstra Masraflar:")
    for amount, code in surcharge_matches:
        if code.upper() in surcharge_keywords:
            extra_total += float(amount)
            print(f"      + {code}: {amount}")

    # B) Ana Navlun Fiyatları (Base Price)
    base_price_20 = 0.0
    base_price_40 = 0.0
    
    # "USD 2800/20" formatını yakala
    p20_match = re.search(r'USD\s*([\d\.]+)\s*/\s*20', body_text, re.IGNORECASE)
    if p20_match: base_price_20 = float(p20_match.group(1))

    # "USD 4300/40" formatını yakala
    p40_match = re.search(r'USD\s*([\d\.]+)\s*/\s*40', body_text, re.IGNORECASE)
    if p40_match: base_price_40 = float(p40_match.group(1))
    
    # 40HC Farkı genelde +100 USD'dir
    base_price_40hc = base_price_40 + 100

    # C) Hedef Limanlar (POD)
    target_pods = ['AMBARLI', 'DERINCE', 'MERSIN', 'ISTANBUL'] # Varsayılan
    # Metindeki "POD: AMBARLI / DERINCE" kısmını bulur
    pod_match = re.search(r'POD:\s*(.*?)(?:\n|\r|USD|20\')', body_text, re.IGNORECASE | re.DOTALL)
    
    if pod_match:
        raw_pods = pod_match.group(1).replace('\n', ' ').split('/')
        found_pods = [p.strip().upper() for p in raw_pods if len(p.strip()) > 2]
        if found_pods: target_pods = found_pods

    print(f"   ℹ️  POD: {target_pods}")
    print(f"   ℹ️  Base Rate: 20'={base_price_20}, 40'={base_price_40}")

    # -------------------------------------------------------------------------
    # ADIM 2: METİN İÇİNDEKİ ANA HATLARI (DIRECT) BULMA
    # -------------------------------------------------------------------------
    # Örnek Metin: "POL: Busan / Qingdao / Shanghai..."
    
    # Regex ile POL kısmını çek (POL: kelimesinden sonra, POD veya USD gelene kadar)
    pol_text_match = re.search(r'POL:\s*(.*?)(?:\n\n|POD:|USD)', body_text, re.IGNORECASE | re.DOTALL)
    
    if pol_text_match:
        # Satır sonlarını boşlukla değiştir ve / işaretine göre böl
        raw_base_pols = pol_text_match.group(1).replace('\n', ' ').split('/')
        base_pols = [p.strip().upper() for p in raw_base_pols if len(p.strip()) > 2]
        
        print(f"   ✅ Metin Bazlı Ana Limanlar (Direct) Bulundu: {len(base_pols)} adet -> {base_pols}")
        
        for pol in base_pols:
            # "LOADING" veya "CODE" gibi gereksiz kelimeleri temizle
            if "CODE" in pol or "LOADING" in pol: continue
            
            for pod in target_pods:
                zim_results.append({
                    'POL': pol,         
                    'VIA': "",  # Ana hatlar DIRECT kabul edilir
                    'POD': pod,             
                    'TOCITY': "", 'CURR': "USD", 'FREETIME': "14", 'TRANSIT': "",
                    'CUSTOMERDESCRIPTION': "", # Main'de temizleniyor
                    'TYPE_1': "20 DC", 'AMOUNT_1': base_price_20 + extra_total,
                    'TYPE_2': "40 DC", 'AMOUNT_2': base_price_40 + extra_total,
                    'TYPE_3': "40 HC", 'AMOUNT_3': base_price_40hc + extra_total
                })
    else:
        print("   ⚠️ Uyarı: Metin içinde 'POL:' anahtar kelimesiyle ana hatlar bulunamadı.")

    # -------------------------------------------------------------------------
    # ADIM 3: TABLO İÇİNDEKİ YAN HATLARI (FEEDER) BULMA
    # -------------------------------------------------------------------------
    if body_html:
        try:
            dfs = pd.read_html(body_html, header=None)
        except: dfs = []

        KEYWORDS_POL = ["POL", "ORIGIN", "LOAD", "PORT OF LOADING"]
        KEYWORDS_VIA = ["VIA", "T/S", "TRANSHIPMENT"]

        for table_idx, df in enumerate(dfs):
            header_row_idx = -1
            is_valid_table = False

            # Tablo başlığını bul
            for i in range(min(10, len(df))):
                row_str = " ".join([str(x).upper() for x in df.iloc[i].values])
                if any(k in row_str for k in KEYWORDS_POL) and any(k in row_str for k in KEYWORDS_VIA):
                    header_row_idx = i
                    is_valid_table = True
                    break
            
            if not is_valid_table: continue

            print(f"   ✅ Tablo #{table_idx+1} (Feeder) işleniyor...")

            current_df = df.copy()
            current_df.columns = current_df.iloc[header_row_idx]
            current_df = current_df.iloc[header_row_idx+1:].reset_index(drop=True)
            current_df.columns = [str(c).strip().upper() for c in current_df.columns]

            col_pol = next((c for c in current_df.columns if any(k in c for k in KEYWORDS_POL)), None)
            col_via = next((c for c in current_df.columns if any(k in c for k in KEYWORDS_VIA)), None)
            
            col_20 = next((c for c in current_df.columns if '20' in c), None)
            col_40 = next((c for c in current_df.columns if '40' in c and 'HQ' not in c and 'HC' not in c), None)
            col_40hq = next((c for c in current_df.columns if 'HQ' in c or 'HC' in c), None)

            if not col_pol or not col_via: continue 

            for idx, row in current_df.iterrows():
                raw_pol = str(row[col_pol]).strip().upper()
                
                # VIA Verisini Çek
                raw_via = ""
                if col_via:
                    val = row[col_via]
                    if pd.notna(val): raw_via = str(val).strip().upper()
                
                if raw_via in ['NAN', 'NONE', 'NAT', 'nan', 'None', '']: continue # Tabloda VIA boşsa alma (Çünkü ana hatları zaten metinden aldık)
                
                # Başlık satırının tekrarıysa atla
                if any(k in raw_pol for k in KEYWORDS_POL): continue 
                if len(raw_pol) < 2 or raw_pol == 'NAN': continue

                list_via = [v.strip() for v in raw_via.split('/') if len(v.strip()) > 1]
                list_pol = [p.strip() for p in raw_pol.split('/') if len(p.strip()) > 1]

                def clean_price(val):
                    try: return float(re.sub(r'[^\d.]', '', str(val).upper().replace('USD','')))
                    except: return 0.0

                f_20 = clean_price(row.get(col_20, 0))
                f_40 = clean_price(row.get(col_40, 0))
                f_40hq = clean_price(row.get(col_40hq, 0))
                if f_40hq == 0 and f_40 > 0: f_40hq = f_40

                for pol in list_pol:
                    for via in list_via:
                        mapped_via = PORT_CODE_MAP.get(via, via)
                        if "DIRECT" in mapped_via: mapped_via = ""

                        # Tablo Fiyatı + Base Fiyat
                        final_20 = base_price_20 + f_20 + extra_total
                        final_40 = base_price_40 + f_40 + extra_total
                        final_40hq = base_price_40hc + f_40hq + extra_total

                        for pod in target_pods:
                            zim_results.append({
                                'POL': pol,         
                                'VIA': mapped_via,        
                                'POD': pod,             
                                'TOCITY': "", 'CURR': "USD", 'FREETIME': "14", 'TRANSIT': "",
                                'CUSTOMERDESCRIPTION': "", 
                                'TYPE_1': "20 DC", 'AMOUNT_1': final_20,
                                'TYPE_2': "40 DC", 'AMOUNT_2': final_40,
                                'TYPE_3': "40 HC", 'AMOUNT_3': final_40hq
                            })

    print(f"✅ İşlem Tamamlandı. Toplam {len(zim_results)} rota oluşturuldu.")
    return zim_results
# =============================================================================
# AYARLAR 
# =============================================================================

GOOGLE_API_KEY = "AIzaSyCP2z4FSIQRDVBXJyiQ9g3B-uYYGyUZHHU" # Key'iniz

BASE_HEADER_COLS = ["POL", "VIA", "POD", "TOCITY", "CURR", "FREETIME", "TRANSIT", "CUSTOMERDESCRIPTION"]

# ETS EUR -> USD Paritesi
EUR_TO_USD = 1.0

# DEBUG MODU
DEBUG_MODE = True

# Türk Limanları (İthalat listesinde POL olmamalılar)
TURKISH_PORTS = [
    "ISTANBUL", "IZMIT", "ALIAGA", "MERSIN", "ISKENDERUN", "GEMLIK", "IZMIR", 
    "ANTALYA", "SAMSUN", "TRABZON", "AMBARLI", "GEBZE", "KOCAELI", "YARIMCA", 
    "TURKEY", "TURKIYE", "TR", "TÜRKIYE", "KUMPORT", "MARPORT", "MARDAS", "DERINCE"
]

# Temizlenecek Ülkeler ve Kelimeler
COUNTRIES_AND_JUNK = [
    "SOUTH KOREA", "KOREA", "CHINA", "VIETNAM", "THAILAND", "JAPAN", "MALAYSIA", "INDONESIA", 
    "INDIA", "PAKISTAN", "BANGLADESH", "SRI LANKA", "PHILIPPINES", "MYANMAR", "CAMBODIA", 
    "EGYPT", "SPAIN", "ITALY", "GERMANY", 
    "FRANCE", "BELGIUM", "NETHERLANDS", "UK", "UNITED KINGDOM", "RUSSIA", "U.A.E",
    "PORT", "TERMINAL", "HARBOUR", "NEW PORT", "OLD PORT", "TP", "TZ", "SK", "BUSAN NEW",
    "SELANO", "GUANGDONG", "SHENZHEN", "ZHEJIANG", "FUJIAN", "MAIN PORTS", "BASE PORTS",
    "REPUBLIC OF", "TAIWAN", 
]

# Çin Limanları
CHINA_PORTS = [
    "SHANGHAI", "NINGBO", "QINGDAO", "XIAMEN", "SHEKOU", "YANTIAN", "NANSHA", 
    "DALIAN", "TIANJIN", "XINGANG", "LIANYUNGANG", "FOSHAN", "HUANGPU", "CHIWAN",
    "FUZHOU", "CHANGZHOU", "CHONGQING", "NANTONG", "NANJING", "ZHANGJIAGANG",
    "SHANSHUI", "SHANTOU", "ZHONGSHAN", "XIAOLAN", "SANSHAN", "JIANGMEN", "JIUJIANG",
    "WUHU", "WUHAN", "WENZHOU", "YANGZHOU", "BEIJIAO", "LEILU", "RONGQI", "CIVET",
    "DOUMEN", "GAOLAN", "ZHAPU", "ZHANJIANG", "ANQING", "HUIZHOU", "JUIJIANG", 
    "QINZHOU" 
]

TERMINAL_TO_CITY_MAP = {
    "KUMPORT": "ISTANBUL", "MARDAS": "ISTANBUL", "MARDAŞ": "ISTANBUL", "MARPORT": "ISTANBUL",
    "AMBARLI": "ISTANBUL", "HAYDARPASA": "ISTANBUL", "HAYDARPAŞA": "ISTANBUL",
    "ISTANBUL": "ISTANBUL", "TURKEY BASE": "ISTANBUL", "TR BASE": "ISTANBUL", "TURKEY": "ISTANBUL",
    "EVYAP": "IZMIT", "DP WORLD": "IZMIT", "DPWORLD": "IZMIT", 
    "YILPORT": "IZMIT", "YARIMCA": "IZMIT", "KOCAELI": "IZMIT", "IZMIT": "IZMIT",
    "BELDEPORT": "IZMIT", "SAFIPORT": "DERINCE", "DERINCE": "DERINCE",
    "RODAPORT": "GEMLIK", "GEMPORT": "GEMLIK", "BORUSAN": "GEMLIK", "GEMLIK": "GEMLIK",
    "NEMPORT": "ALIAGA", "TCEGE": "ALIAGA", "SOCAR": "ALIAGA", "ALIAGA": "ALIAGA", 
    "ALSANCAK": "IZMIR", "IZMIR": "IZMIR",
    "MERSIN": "MERSIN", "MIP": "MERSIN",
    "ISKENDERUN": "ISKENDERUN", "LIMAK": "ISKENDERUN",
    "SAMSUN": "SAMSUN", "TRABZON": "TRABZON", "ANTALYA": "ANTALYA"
}

EXCEL_SHEET_BLACKLIST = ["TAO", "TAD", "SURCHARGE", "ALLOCATION", "LOCAL", "DEMURRAGE", "FREE TIME", "ADDITIONALS"]

# GLOBAL CONTEXTLER
GLOBAL_ONE_CONTEXT = {
    "is_active": False,
    "obs": 0, "pss": 0, "est": 0
}

GLOBAL_YANGMING_CONTEXT = {
    "is_active": False,
    "isps": 0,      # USD
    "pss": 0,       # USD 
    "ets_20": 0,    # EUR (1:1 eklenecek)
    "ets_40": 0,    # EUR (1:1 eklenecek)
    "ets_exclude_china": False,
    "is_ets_eur": False
}

# MSC CONTEXT
GLOBAL_MSC_CONTEXT = {
    "is_active": False,
    "cls": 0,           
    "crs_china": 0,    
    "crs_other": 0,    
    "cdd": 0            
}

# ZIM CONTEXT
GLOBAL_ZIM_CONTEXT = {
    "is_active": False,
    "extra_total": 0    # Toplam eklenecek tutar (Örn: 10 IPS + 30 BL = 40)
}

# SEA LEAD CONTEXT
GLOBAL_SEALEAD_CONTEXT = {
    "is_active": False,
    "eca": 0            # Per TEU (20'lik x1, 40'lık x2)
}

# COSCO CONTEXT 
GLOBAL_COSCO_CONTEXT = {
    "is_active": False,
    "ets_eur": 0 # Per TEU (20'lik x1, 40'lık x2)
}

# =============================================================================
# MODEL YÖNETİCİSİ
# =============================================================================

class ModelManager:
    def __init__(self, api_key):
        self.api_key = api_key
        self.available_models = []
        self.current_index = 0
        self.refresh_models()

    def refresh_models(self):
        print("  🌐 Google AI Modelleri Taranıyor...", flush=True)
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={self.api_key}"
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                data = r.json()
                raw_models = [m['name'].replace('models/', '') for m in data.get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
                flash_models = [m for m in raw_models if 'flash' in m]
                pro_models = [m for m in raw_models if 'pro' in m and 'flash' not in m]
                others = [m for m in raw_models if m not in flash_models and m not in pro_models]
                flash_models.sort(reverse=True) 
                pro_models.sort(reverse=True)
                self.available_models = flash_models + pro_models + others
                if not self.available_models: self.available_models = ["gemini-1.5-flash", "gemini-1.5-pro"]
                print(f"  ✅ Modeller: {self.available_models}", flush=True)
            else:
                self.available_models = ["gemini-1.5-flash", "gemini-1.5-pro"]
        except:
            self.available_models = ["gemini-1.5-flash", "gemini-1.5-pro"]

    def get_current_model(self):
        return self.available_models[self.current_index]

    def switch_to_next_model(self):
        old_model = self.available_models[self.current_index]
        self.current_index = (self.current_index + 1) % len(self.available_models)
        print(f"      🔄 Model Değiştiriliyor: {old_model} ➡️ {self.available_models[self.current_index]}", flush=True)

model_mgr = ModelManager(GOOGLE_API_KEY)

# =============================================================================
# YARDIMCI FONKSİYONLAR
# =============================================================================

def global_clean_port_name(text):
    if not isinstance(text, str): return ""
    text = text.upper().replace("İ", "I").strip()
    
    if text == "NGB": return "NINGBO"
    
    text = re.sub(r'(\))([A-Z])', r'\1 / \2', text)
    text = re.sub(r'\b(PORT|TERMINAL)\b', '', text)
    for junk in COUNTRIES_AND_JUNK:
        pattern = r'\b' + re.escape(junk) + r'\b'
        text = re.sub(pattern, "", text)
    text = text.replace(",", "").replace(".", "").replace("-", " ")
    text = " ".join(text.split())
    return text.strip()

# =============================================================================
# SON KONTROL VE TEMİZLİK FONKSİYONLARI (MSC & ONE)
# =============================================================================

# MSC çıktısında silinecek liman kodları
MSC_YASAKLI_KODLAR = [
    "CNSNH", "CNSGB", "CNQIN", "CNSAD", "CNTXG", "CNYTN", "CNSHK", 
    "CNXNG", "CNXING", "CNDAL", "CNXIA", "CNQNQ", "CNOHG", "CNCHZ", 
    "CNOCQ", "CNFUZ", "CNJIC", "CNJJG", "CNLYU", "CNNHJ", "CNNNJ", 
    "CNNTG", "CNZJG", "CNZAP", "CNCHE", "CNSHS", "CNTAG", "CNTZO", 
    "CNWZU", "CNWUH", "CNWUU", "CNYAZ", "CNYIH", "CNYEY", "CNZEN", 
    "CNWAZ", "CNLUZ", "CNYND", "CNDAF", "CNQIZ", "CNGAO", "CNJJA", 
    "CNSHU", "CNIHZ", "CNZHQ", "CNLEL", "CNHAU", "CNHUA", "CNJIA", 
    "CNNSJ", "CNRQI", "CNBJO", "CNSHN", "CNXIO", "CNZHJ", "CNZSH", 
    "CNZHU", "CNOZX", "CNWUZ", "CNHKG",
    "KRPUS", "VNTCT", "VNHCM", "THLEM", "HKHKG", "SGSIN", "TWKLG", 
    "TWKSG", "TWTXG", "TWTYG", "JPTYO", "JPNGO", "JPYOK", "JPKBE", 
    "JPOSA", "IDJKT", "IDSUB", "VNHAI", "VNDAD", "THBKK", "THLKR", 
    "MYPKL", "PHZMP", "PHMNL", "KHPKH", "KHOUX", "CNNSA", "CNXMN",
    "CNFOC", "CNSWA", "CNQZH", "CNSHA", "CNNGB", "CNTAO", "CNXGG",
    "CNDLC", "MYPGU", "MYPEN", "MYPKG", "MYTPP", "IDBLW", "JAVA",
    "IDSRG", "IDPNJ", "THLCH", "VNCLN", "VNVUT", "VNUIH", "KRKAN",
    "KRINC", "TWKHH", "PHMNN", "JPHKT", "JPUKB", "JPYKK"
]

def apply_final_corrections(data, filename):
    """
    Veriler CSV'ye yazılmadan önce firmaya özel son temizlikleri yapar.
    data: List of dicts (Tablo satırları)
    filename: İşlenen dosyanın adı
    """
    if not data: return data
    
    fname = filename.upper()
    
    # --- 1. MSC İÇİN POL TEMİZLİĞİ ---
    if "MSC" in fname:
        for row in data:
            if row.get('POL'):
                pol_val = str(row['POL'])
                # Listedeki her kodu tek tek sil
                for bad_code in MSC_YASAKLI_KODLAR:
                    # Kodu silerken boşluk bırakmamaya çalışalım
                    pol_val = pol_val.replace(bad_code, "")
                
                # Gereksiz boşlukları ve karakterleri temizle
                row['POL'] = pol_val.strip().replace("  ", " ")

    # --- 2. ONE LINE İÇİN POD TEMİZLİĞİ (FAK SİLME) ---
    if "ONE" in fname:
        for row in data:
            if row.get('POD'):
                pod_val = str(row['POD'])
                pod_val = re.sub(r'\bFAK\b', '', pod_val, flags=re.IGNORECASE)
                row['POD'] = pod_val.strip()

    # --- 3. ZIM İÇİN CUSTOMER DESCRIPTION TEMİZLİĞİ ---
    if "ZIM" in fname:
        for row in data:
            # Bu sütunu zorla boşalt
            row['CUSTOMERDESCRIPTION'] = ""
                
    return data

def global_map_pod(pod_name, target_pods_context=None):
    clean = global_clean_port_name(pod_name)
    if target_pods_context and len(target_pods_context) == 1 and "ALIAGA" in target_pods_context:
        if clean in ["IZMIR", "IZT", "ALSANCAK"]:
            return "ALIAGA"
    if clean in TERMINAL_TO_CITY_MAP: return TERMINAL_TO_CITY_MAP[clean]
    for key, val in TERMINAL_TO_CITY_MAP.items():
        if key in clean: return val
    return clean

def is_turkish_port(name):
    n = global_map_pod(name)
    if n in TERMINAL_TO_CITY_MAP.values(): return True
    if name.upper() in TURKISH_PORTS: return True
    return False

def extract_target_pods_from_text(text):
    text_upper = text[:6000].upper() 
    found_pods = set()
    
    pod_match = re.search(r'POD\s*[:]\s*(.*?)[\r\n]', text_upper)
    if pod_match:
        pod_line = pod_match.group(1)
        parts = re.split(r'[/\-,\s]+', pod_line)
        for part in parts:
            clean_part = part.strip()
            if len(clean_part) < 3: continue
            if clean_part in TERMINAL_TO_CITY_MAP:
                found_pods.add(TERMINAL_TO_CITY_MAP[clean_part])
            elif clean_part in ["ISTANBUL", "DERINCE", "IZMIT", "MERSIN", "IZMIR", "ALIAGA", "ISKENDERUN", "AMBARLI"]:
                if clean_part == "AMBARLI": found_pods.add("ISTANBUL")
                else: found_pods.add(clean_part)
                
    if not found_pods:
        for terminal, city in TERMINAL_TO_CITY_MAP.items():
            if terminal in text_upper:
                found_pods.add(city)
                
    if found_pods:
        if "ALIAGA" in found_pods and "IZMIR" in found_pods:
            found_pods.remove("IZMIR")
        return list(found_pods)

    if "TURKEY" in text_upper or "TR BASE" in text_upper:
         return ["ISTANBUL", "IZMIT", "ALIAGA", "MERSIN"]
         
    return ["ISTANBUL"]

def isolate_latest_email_body(text):
    if not text: return ""
    separators = [
        r'(?m)^From:\s.*Sent:\s',
        r'(?m)^Kimden:\s.*Gönderilen:\s',
        r'(?m)^-{3,}\s?Original\s+Message\s?-{3,}',
        r'(?m)^_{10,}',
        r'(?m)^On\s+.*,\s+.*wrote:$',
        r'(?m)^On\s+.*\s+wrote:$'
    ]
    best_cut_index = len(text)
    for pat in separators:
        match = re.search(pat, text, re.IGNORECASE | re.DOTALL)
        if match:
            if match.start() < best_cut_index:
                best_cut_index = match.start()
    
    if best_cut_index < len(text):
        return text[:best_cut_index]
    return text

def safe_float(val):
    try:
        if val is None: return 0.0
        return float(val)
    except:
        return 0.0

# --- MASRAF YAKALAYICILAR ---
def extract_one_surcharges(text):
    text_upper = text[:4000].upper()
    surcharges = {"obs": 0, "pss": 0, "est": 0}
    obs_match = re.search(r'OBS.*?(\d{1,4})', text_upper)
    if obs_match: surcharges["obs"] = safe_float(obs_match.group(1))
    pss_match = re.search(r'PSS.*?(\d{1,4})', text_upper)
    if pss_match: surcharges["pss"] = safe_float(pss_match.group(1))
    est_match = re.search(r'EST.*?(\d{1,4})', text_upper)
    if est_match: surcharges["est"] = safe_float(est_match.group(1))
    return surcharges

def extract_yangming_surcharges(text):
    text_upper = text.upper()
    surcharges = {"isps": 0, "pss": 0, "ets_20": 0, "ets_40": 0, "ets_exclude_china": False, "is_ets_eur": False}
    scan_limit = text_upper 

    if "ISPS" in scan_limit and ("DAHİL DEĞİL" in scan_limit or "HARİÇ" in scan_limit):
        isps_match = re.search(r'ISPS.*?(\d{1,3})', scan_limit)
        if isps_match: surcharges["isps"] = safe_float(isps_match.group(1))

    if "PSS" in scan_limit or "PEAK SEASON" in scan_limit or "GRI" in scan_limit:
        pss_match = re.search(r'(?:PSS|GRI).*?(\d{2,4})', scan_limit)
        if pss_match: surcharges["pss"] = safe_float(pss_match.group(1))

    if "ETS" in scan_limit or "EMISSION" in scan_limit or "EU ETS" in scan_limit:
        if "EUR" in scan_limit: surcharges["is_ets_eur"] = True
        ets20_match = re.search(r'(?:EUR|USD)\s*(\d{1,3})\s*[\/\\]?20', scan_limit)
        if ets20_match: surcharges["ets_20"] = safe_float(ets20_match.group(1))
        ets40_match = re.search(r'(?:EUR|USD)\s*(\d{1,3})\s*[\/\\]?40', scan_limit)
        if ets40_match: surcharges["ets_40"] = safe_float(ets40_match.group(1))
        
        if "ÇİN" in scan_limit and ("HARİÇ" in scan_limit or "EXCEPT" in scan_limit):
            surcharges["ets_exclude_china"] = True
            
    if DEBUG_MODE:
        print(f"    🔎 [DEBUG - MASRAFLAR] ISPS: {surcharges['isps']} | PSS: {surcharges['pss']} | ETS20: {surcharges['ets_20']} | ÇİN HARİÇ: {surcharges['ets_exclude_china']}", flush=True)

    return surcharges

def extract_msc_surcharges(text):
    text_upper = text.upper()
    surcharges = {"cls": 0, "crs_china": 0, "crs_other": 0, "cdd": 0}
    print(f"    🔍 [MSC] Mail içerisinde ek masraflar (CLS, CRS, CDD) taranıyor...", flush=True)
    cls_match = re.search(r'CLS\s*[-–:]?\s*(?:USD)?\s*(\d{1,4})', text_upper)
    if cls_match: surcharges["cls"] = safe_float(cls_match.group(1))
    crs_eur_match = re.search(r'CRS.*?EUR\s*(\d{1,4})', text_upper)
    if crs_eur_match: surcharges["crs_china"] = safe_float(crs_eur_match.group(1))
    crs_usd_match = re.search(r'CRS.*?USD\s*(\d{1,4})', text_upper)
    if crs_usd_match: surcharges["crs_other"] = safe_float(crs_usd_match.group(1))
    if not crs_eur_match and crs_usd_match: surcharges["crs_china"] = surcharges["crs_other"]
    cdd_match = re.search(r'CDD\s*[-–:]?\s*(?:USD)?\s*(\d{1,4})', text_upper)
    if cdd_match: surcharges["cdd"] = safe_float(cdd_match.group(1))
    
    if any(v > 0 for v in surcharges.values()):
        print(f"    ✅ [MSC] MASRAFLAR BULUNDU: CLS={surcharges['cls']} | CRS(CN)={surcharges['crs_china']} | CRS(Other)={surcharges['crs_other']} | CDD={surcharges['cdd']}", flush=True)
    else:
        print(f"    ⚠️ [MSC] Mailde belirtilen bir ek masraf bulunamadı. (Fiyatlara 0 eklenecek)", flush=True)
    return surcharges

def extract_zim_surcharges(text):
    text_lower = text.lower()
    total_extra = 0
    subj_index = text_lower.find("subj to")
    if subj_index == -1: subj_index = text_lower.find("subject to")
    if subj_index != -1:
        search_area = text_lower[subj_index:subj_index+250]
        matches = re.findall(r'usd\s?(\d+)', search_area)
        for m in matches:
            try: total_extra += safe_float(m)
            except: pass
    if DEBUG_MODE: print(f"    🔎 [ZIM DEBUG] Bulunan Ekstra Masraf Toplamı: {total_extra} USD", flush=True)
    return {"extra_total": total_extra}

def extract_sealead_surcharges(text):
    text_upper = text.upper()
    surcharges = {"eca": 0}
    match = re.search(r'ECA.*?USD\s*(\d{1,3})', text_upper)
    if match:
        surcharges["eca"] = safe_float(match.group(1))
        print(f"    ✅ [SEA LEAD] ECA BULUNDU: {surcharges['eca']} USD/TEU", flush=True)
    else:
        print(f"    ⚠️ [SEA LEAD] ECA Bulunamadı (0 kabul edilecek).", flush=True)
    return surcharges

def extract_cosco_surcharges(text):
    """
    COSCO için ETS masrafını tarar.
    Örn: "ETS ücretimiz var 57eur/teu"
    """
    text_upper = text.upper()
    surcharges = {"ets_eur": 0}
    
    # ETS + rakam + EUR/TEU yapısını ara
    match = re.search(r'ETS.*?(\d{1,3})\s*EUR', text_upper)
    
    if match:
        surcharges["ets_eur"] = safe_float(match.group(1))
        print(f"    ✅ [COSCO] ETS BULUNDU: {surcharges['ets_eur']} EUR/TEU", flush=True)
    else:
        print(f"    ⚠️ [COSCO] ETS Bulunamadı (0 kabul edilecek).", flush=True)
        
    return surcharges

# =============================================================================
# DOSYA EKİ ÇIKARMA
# =============================================================================

def save_excel_attachments_with_prefix(msg_obj, file_type, prefix=""):
    extracted = []
    if file_type == "eml":
        for part in msg_obj.walk():
            if part.get_content_disposition() == 'attachment':
                fname = part.get_filename()
                if fname and (fname.lower().endswith(('.xlsx', '.xls', '.pdf'))): # PDF eklendi
                    safe_name = re.sub(r'[\\/*?:"<>|]', "", fname)
                    final_name = f"{prefix}{safe_name}"
                    with open(final_name, 'wb') as f_out:
                        f_out.write(part.get_payload(decode=True))
                    print(f"    📎 [EK ÇIKARILDI] {final_name}", flush=True)
                    extracted.append(final_name)
    elif file_type == "msg":
        for att in msg_obj.attachments:
            if att.longFilename and (att.longFilename.lower().endswith(('.xlsx', '.xls', '.pdf'))): # PDF eklendi
                safe_name = re.sub(r'[\\/*?:"<>|]', "", att.longFilename)
                final_name = f"{prefix}{safe_name}"
                with open(final_name, 'wb') as f:
                    f.write(att.data)
                print(f"    📎 [EK ÇIKARILDI] {final_name}", flush=True)
                extracted.append(final_name)
    return extracted

# =============================================================================
# PDF İŞLEME
# =============================================================================

def pdf_extract_text(pdf_path):
    full_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(layout=True)
                if text: full_text += text + "\n"
    except: return ""
    return full_text





def process_pdf_file(pdf_path):
    print(f"  > PDF Analiz Ediliyor: {os.path.basename(pdf_path)}", flush=True)
    text = pdf_extract_text(pdf_path) 
    if not text: return None
    
    is_cosco = "COSCO" in os.path.basename(pdf_path).upper() or "COSCO" in text.upper()
    
    # ----------------------------------------------------------------------------------------------
    # COSCO PDF İŞLEME MANTIĞI
    # ----------------------------------------------------------------------------------------------
    if is_cosco:
        print("    ℹ️ COSCO PDF Tespit Edildi...", flush=True)
        
        # 1. ETS Eklemesi
        ets_per_teu = 0
        if GLOBAL_COSCO_CONTEXT["is_active"]:
            ets_per_teu = GLOBAL_COSCO_CONTEXT["ets_eur"] * EUR_TO_USD
        
        # 2. Add-on Tablosunu Bulma
        add_on_rates = {}
        lines = text.split('\n')
        
        target_addon_ports = ["GEMLIK", "IZMIR", "MERSIN", "ALIAGA", "ISKENDERUN", "SAMSUN"]
        print(f"    🔎 [COSCO] Add-on tablosu taranıyor (Hedefler: {target_addon_ports})...", flush=True)
        
        for line in lines:
            line_upper = line.upper()
            line_upper = line_upper.replace('İ', 'I').replace('Ğ', 'G').replace('Ü', 'U').replace('Ş', 'S').replace('Ö', 'O').replace('Ç', 'C')
            
            for t_port in target_addon_ports:
                if t_port in line_upper:
                    price_match = re.search(r'(?:[\+]|[:]|USD)?\s*(\d{2,4})', line_upper)
                    if price_match:
                         port_idx = line_upper.find(t_port)
                         after_port = line_upper[port_idx:]
                         num_match = re.search(r'(\d{2,4})', after_port)
                         if num_match:
                             try:
                                 price = float(num_match.group(1))
                                 if price < 1000:
                                     add_on_rates[t_port] = price
                             except: pass

        if add_on_rates:
            print(f"    ➕ [COSCO] Tespit Edilen Add-onlar: {add_on_rates}", flush=True)
        else:
             print(f"    ⚠️ [COSCO] Add-on tablosu bulunamadı, sadece ana limanlar (IST/DER) basılacak.", flush=True)

        # 3. Ana Liman Fiyatlarını Okuma
        main_rates = []
        for line in lines:
            line_upper = line.upper()
            line_upper = line_upper.replace('İ', 'I')
            
            numbers = re.findall(r'\b\d{3,4}\b', line_upper)
            if len(numbers) >= 2: 
                try:
                    first_num_idx = line_upper.find(numbers[0])
                    pol_raw = line_upper[:first_num_idx].strip()
                    pol_clean = global_clean_port_name(pol_raw)
                    
                    if len(pol_clean) > 2 and not is_turkish_port(pol_clean):
                        p20 = float(numbers[0])
                        p40 = float(numbers[1])
                        p40hc = p40 
                        if len(numbers) >= 3: p40hc = float(numbers[2])
                        
                        main_rates.append({"pol": pol_clean, "p20": p20, "p40": p40, "p40hc": p40hc})
                except: continue

        # 4. Rotaları Oluşturma
        default_pods = ["ISTANBUL", "DERINCE"] 
        all_rows = []
        
        for rate in main_rates:
            base_20 = rate["p20"]
            base_40 = rate["p40"]
            base_40hc = rate["p40hc"]
            pol_current = rate["pol"]
            
            applied_ets_unit = 0 
            if is_cosco and GLOBAL_COSCO_CONTEXT["is_active"]:
                if not is_china_region_port(pol_current):
                    applied_ets_unit = ets_per_teu
                else:
                    applied_ets_unit = 0 
            
            final_base_20 = base_20 + applied_ets_unit
            final_base_40 = base_40 + (applied_ets_unit * 2)
            final_base_40hc = base_40hc + (applied_ets_unit * 2)
            
            for d_pod in default_pods:
                all_rows.append({
                    "POL": pol_current, "VIA": "", "POD": d_pod, "TOCITY": "", "CURR": "USD", "FREETIME": "",
                    "TYPE_1": "20 DC", "AMOUNT_1": final_base_20,
                    "TYPE_2": "40 DC", "AMOUNT_2": final_base_40,
                    "TYPE_3": "40 HC", "AMOUNT_3": final_base_40hc
                })
            
            if add_on_rates:
                 for addon_pod, extra_cost in add_on_rates.items():
                    pod_mapped = global_map_pod(addon_pod)
                    addon_20 = final_base_20 + extra_cost
                    addon_40 = final_base_40 + extra_cost 
                    addon_40hc = final_base_40hc + extra_cost
                    
                    all_rows.append({
                        "POL": pol_current, "VIA": "", "POD": pod_mapped, "TOCITY": "", "CURR": "USD", "FREETIME": "",
                        "TYPE_1": "20 DC", "AMOUNT_1": addon_20,
                        "TYPE_2": "40 DC", "AMOUNT_2": addon_40,
                        "TYPE_3": "40 HC", "AMOUNT_3": addon_40hc
                    })

        # =========================================================================
        # 5.POL KONTROL VE FİLTRELEME
        # =========================================================================
        
        # Filtrelenecek (Yasaklı) Limanlar Listesi - Büyük harf ve İngilizce karakter ile
        forbidden_pols = ["ALIAGA", "IZMIR", "GEMLIK", "ISKENDERUN", "MERSIN"]
        
        filtered_rows = []
        dropped_count = 0

        for row in all_rows:
            # POL verisini al, büyük harfe çevir ve Türkçe karakterleri temizle
            current_pol = row.get("POL", "").upper()
            current_pol = current_pol.replace('İ', 'I').replace('Ğ', 'G').replace('Ü', 'U').replace('Ş', 'S').replace('Ö', 'O').replace('Ç', 'C')
            
            # Yasaklı kelimelerden herhangi biri POL içinde geçiyor mu?
            is_forbidden = any(f_port in current_pol for f_port in forbidden_pols)
            
            if is_forbidden:
                dropped_count += 1
                # LOG
                # print(f"    🗑️ [COSCO] Silinen Satır (POL Yasaklı): {row['POL']}", flush=True)
            else:
                filtered_rows.append(row)
        
        if dropped_count > 0:
            print(f"    🧹 [COSCO] Temizlik: {dropped_count} adet Türk limanı çıkışlı (POL) satır silindi.", flush=True)

        return filtered_rows

    return None



# =============================================================================
# MSG / EML İŞLEME
# =============================================================================

def msg_query_gemini(chunk_text, chunk_index, total_chunks, is_oocl=False, is_sealead=False):
    if is_oocl:
        prompt = """
        Extract ALL ocean freight rates (Main Ports AND Feeder/Barge Ports).
        CRITICAL FOR OOCL:
        1. Capture "Main Port" rates (e.g. Shanghai: 2000).
        2. Capture "Feeder/Add-on" rates (e.g. Wenzhou via Shanghai: +150).
        3. EXTRACT THE 'VIA' PORT CORRECTLY (e.g. if text says "via Shanghai", extract "Shanghai").
        4. IGNORE "Inland", "Truck", "Rail" or "CNY" rates.
        JSON Format: { "routes": [ {"pol": "Wenzhou", "via": "Shanghai", "pod": "Istanbul", "curr": "USD", "p20": 150, "p40": 250, "p40hc": 250} ] }
        """
    elif is_sealead:
        prompt = """
        Extract ocean freight rates.
        SEA LEAD emails often list multiple POLs and PODs followed by a single price block.
        INSTRUCTION:
        If you see a list of origins (POLs) and destinations (PODs) associated with one price set, 
        GENERATE A ROUTE FOR EACH COMBINATION.
        OUTPUT JSON: { "routes": [ {"pol": "Qingdao", "via": "", "pod": "Ambarli", "curr": "USD", "p20": 4250, "p40": 6100, "p40hc": 6100}, ... ] }
        """
    else:
        prompt = """
        Extract ocean freight rates.
        Target: POL, POD, PRICE (20, 40, 40HC).
        CRITICAL INSTRUCTIONS:
        1. Extract ONLY Import rates TO Turkey.
        2. If POL contains slashes, keep it as is.
        3. If multiple ports listed, extract as is.
        OUTPUT JSON FORMAT: { "routes": [ {"pol": "X", "via": "Y", "pod": "Z", "curr": "USD", "p20": 0, "p40": 0, "p40hc": 0} ] }
        """
        
    payload = {"contents": [{"parts": [{"text": prompt + "\n\nTEXT:\n" + chunk_text}]}], "generationConfig": {"responseMimeType": "application/json"}}
    for attempt in range(1, 5):
        current_model = model_mgr.get_current_model()
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{current_model}:generateContent?key={GOOGLE_API_KEY}"
        try:
            r = requests.post(url, json=payload, headers={'Content-Type': 'application/json'}, timeout=60)
            if r.status_code == 200:
                try: return json.loads(r.json()['candidates'][0]['content']['parts'][0]['text'])
                except: model_mgr.switch_to_next_model(); continue
            elif r.status_code in [429, 503, 500]:
                model_mgr.switch_to_next_model(); time.sleep(5); continue
            else: time.sleep(5)
        except: time.sleep(5)
    return {"routes": []}




# Yardımcı fonksiyon: İsim temizleme (Virgül ve parantez temizliği)
def normalize_port_name(port_name):
    if not port_name: return ""
    # "Ningbo, Ningbo, Zhejiang, China" -> "Ningbo" yapmak için virgülden öncesini al
    port_name = port_name.split(',')[0]
    cleaned = port_name.split('(')[0]
    cleaned = re.sub(r'[^A-Z\s]', '', cleaned.upper()).strip()
    return cleaned

def process_msg_file(file_path):
    filename = os.path.basename(file_path).upper()
    if "HAPAG" in filename:
        return None
    print(f"  > MSG/EML İşleniyor: {filename}", flush=True)

    final_text = ""
    msg_obj = None
    file_type = "" 

    if file_path.lower().endswith(".eml"):
        file_type = "eml"
        try:
            with open(file_path, 'rb') as f:
                msg_obj = BytesParser(policy=policy.default).parse(f)
                body = msg_obj.get_body(preferencelist=('plain')).get_content()
                if body: final_text = body
        except Exception as e: return None
    else:
        file_type = "msg"
        try: 
            msg_obj = extract_msg.Message(file_path)
            text_body = msg_obj.body
            if text_body: final_text = text_body
        except: return None
    
    final_text = re.sub(r'\s+', ' ', final_text)
    file_prefix = ""
    
    # --- MARKALAR VE MASRAF ÇEKME İŞLEMLERİ  ---
    is_sealead_file = False
    is_ym_file = False
    if "ONE" in filename or "ONE LINE" in final_text.upper():
        file_prefix = "ONE_LINE_" 
        extras = extract_one_surcharges(final_text)
        if any(extras.values()): 
            GLOBAL_ONE_CONTEXT["is_active"] = True
            GLOBAL_ONE_CONTEXT.update(extras)
           
    elif "YANG MING" in filename or "YANG MING" in final_text.upper():
        print("    ℹ️ YANG MING Tespit Edildi... Geçmiş temizleniyor...", flush=True)
        file_prefix = "YANG_MING_"
        try: final_text = isolate_latest_email_body(final_text)
        except: pass
        
        # Kesme mantığı...
        lower_text = final_text.lower()
        search_start_limit = 150; cut_candidates = []
        if lower_text.find("from:", search_start_limit) != -1: cut_candidates.append(lower_text.find("from:", search_start_limit))
        if lower_text.find("kimden:", search_start_limit) != -1: cut_candidates.append(lower_text.find("kimden:", search_start_limit))
        if lower_text.find("original message", search_start_limit) != -1: cut_candidates.append(lower_text.find("original message", search_start_limit))
        if cut_candidates: final_text = final_text[:min(cut_candidates)]
        
        extras = extract_yangming_surcharges(final_text)
        if extras["isps"] > 0 or extras["ets_20"] > 0: 
            GLOBAL_YANGMING_CONTEXT["is_active"] = True
            GLOBAL_YANGMING_CONTEXT.update(extras)
            
    #  MSC MASRAFLARINI ÇEK ---
    elif "MSC" in filename or "MSC" in final_text.upper():
        file_prefix = "MSC_"
        extras = extract_msc_surcharges(final_text) 
        GLOBAL_MSC_CONTEXT["is_active"] = True
        GLOBAL_MSC_CONTEXT.update(extras)          
        
    elif "ZIM" in filename or "ZIM" in final_text.upper():
        file_prefix = "ZIM_"
        GLOBAL_ZIM_CONTEXT["is_active"] = True
       

    # ---  SEA LEAD MASRAFLARINI ÇEK ---
    elif ("SEA" in filename and "LEAD" in filename) or "SEA-LEAD" in filename:
        file_prefix = "SEALEAD_"
        is_sealead_file = True
        extras = extract_sealead_surcharges(final_text)
        GLOBAL_SEALEAD_CONTEXT["is_active"] = True
        GLOBAL_SEALEAD_CONTEXT.update(extras)           

    # ---  COSCO MASRAFLARINI ÇEK ---
    elif "COSCO" in filename:
        file_prefix = "COSCO_"
        extras = extract_cosco_surcharges(final_text)  
        GLOBAL_COSCO_CONTEXT["is_active"] = True
        GLOBAL_COSCO_CONTEXT.update(extras)            

    elif "OOCL" in filename: 
        file_prefix = "OOCL_"

    extracted_files = []
    if msg_obj:
        extracted_files = save_excel_attachments_with_prefix(msg_obj, file_type, prefix=file_prefix)
        if file_type == "msg": msg_obj.close()

    # Eğer ek varsa çık 
    
    if len(extracted_files) > 0: return None
    
    if len(final_text) < 50: return None
    
    target_pods = extract_target_pods_from_text(final_text)
    
       
    CHUNK_SIZE = 13000; OVERLAP = 1500 
    chunks = []; start = 0
    while start < len(final_text):
        end = min(start + CHUNK_SIZE, len(final_text))
        chunks.append(final_text[start:end])
        if end == len(final_text): break
        start += (CHUNK_SIZE - OVERLAP)
    
    raw_extracted_routes = []
    seen_routes = set()

    is_oocl_file = "OOCL" in filename
    is_ym_file = "YANG" in filename or "YANG" in file_prefix
    
    
   
    OOCL_KNOWN_BASES = ["SHANGHAI", "NINGBO", "QINGDAO", "XIAMEN", "SHEKOU", "YANTIAN", "NANSHA", "SINGAPORE", "HONG KONG", "BUSAN", "PORT KLANG", "DALIAN", "XINGANG", "TIANJIN", "KAOHSIUNG", "YOKOHAMA", "OSAKA", "KOBE", "TOKYO", "HO CHI MINH", "CAT LAI", "HAIPHONG", "LAEM CHABANG", "JAKARTA"]

    for i, chunk in enumerate(chunks):
        res = msg_query_gemini(chunk, i, len(chunks), is_oocl=is_oocl_file, is_sealead=is_sealead_file)
        chunk_has_routes = False

        if res and "routes" in res:
            for r in res["routes"]:
               
                if is_oocl_file:
                    curr = str(r.get("curr", "USD")).upper()
                    if "CNY" in curr or "RMB" in curr: continue
                
                raw_pol_str = r.get("pol", "")
                raw_via = r.get("via", "")
                
                clean_via = global_clean_port_name(raw_via)
                if clean_via == "NGB": clean_via = "NINGBO"

                #  POL verisini temizle ve ayır (Split)
                # Örn: "SHANGHAI / NINGBO" -> ["SHANGHAI", "NINGBO"] olarak listeye çevirir.
                clean_pol_base = global_clean_port_name(raw_pol_str)
                # Sadece / değil , ve ; ile de ayırır.
                pol_parts = [p.strip() for p in re.split(r'[/,;]', clean_pol_base) if len(p.strip()) > 1]
                
                #  Eğer hiç POL bulunamadıysa (Boşsa), bu satırı atla.
                # Yang Ming çıktısındaki boş satırları bu engeller.
                if not pol_parts:
                    continue

                # Ayrıştırılan her bir liman için ayrı satır oluştur
                for single_pol in pol_parts:
                    clean_pol = global_clean_port_name(single_pol)
                    
                    # Türk limanı çıkışlı ise atla (Biz ithalat yapıyoruz)
                    if is_turkish_port(clean_pol): continue
                    
                    clean_pod = global_clean_port_name(r.get("pod", ""))

                    # OOCL SÜTUN KAYMASI DÜZELTMESİ (Mevcut mantık korundu)
                    if is_oocl_file and not clean_via and clean_pod:
                        pod_norm = normalize_port_name(clean_pod)
                        pol_norm = normalize_port_name(clean_pol)
                        
                        OOCL_KNOWN_BASES = ["SHANGHAI", "NINGBO", "QINGDAO", "XIAMEN", "SHEKOU", "YANTIAN", "NANSHA", "SINGAPORE", "HONG KONG", "BUSAN", "PORT KLANG", "DALIAN", "XINGANG", "TIANJIN", "KAOHSIUNG", "YOKOHAMA", "OSAKA", "KOBE", "TOKYO", "HO CHI MINH", "CAT LAI", "HAIPHONG", "LAEM CHABANG", "JAKARTA"]
                        
                        is_pod_base = any(b in pod_norm for b in OOCL_KNOWN_BASES)
                        is_pol_base = any(b in pol_norm for b in OOCL_KNOWN_BASES)
                        
                        if is_pod_base and not is_pol_base:
                            clean_via = clean_pod 
                            clean_pod = ""

                    if clean_pod:
                        clean_pod = global_map_pod(clean_pod, target_pods_context=target_pods)

                    p20_val = safe_float(r.get("p20", 0))

                    new_r = r.copy()
                    
                    # Artık 'clean_pol_base' (hepsi) değil, 'single_pol' (tekil) atanıyor.
                    new_r["pol"] = clean_pol 
                    
                    new_r["via"] = clean_via
                    new_r["pod"] = clean_pod 
                    new_r["p20"] = p20_val 
                    new_r["p40"] = safe_float(r.get("p40", 0))
                    new_r["p40hc"] = safe_float(r.get("p40hc", 0))
                    
                    # Anahtarı da tekil limana göre oluşturuyoruz
                    route_key = (clean_pol, clean_via, clean_pod, p20_val)
                    
                    if route_key not in seen_routes:
                        seen_routes.add(route_key)
                        if len(new_r["pol"]) > 1:
                            raw_extracted_routes.append(new_r)
                            chunk_has_routes = True

    # --- ROTA FİNALİZASYONU (OOCL / STANDART) ---
    final_routes = []
    if is_oocl_file:
        print("    ℹ️ OOCL Modu: Genişletilmiş ve Düzeltilmiş Mantık...", flush=True)
        main_port_rates = {}
        
        # 1. ADIM: Ana Limanları Topla
        for r in raw_extracted_routes:
            p20 = r["p20"] 
            pol = r.get("pol", "").upper()
            pol_normalized = normalize_port_name(pol)
            
            via_raw = r.get("via", "")
            via_normalized = normalize_port_name(via_raw)

            # Fiyat ve Rota Güvenilirlik Kontrolü
            is_high_price = p20 > 1500 
            is_known_base = any(base == pol_normalized for base in OOCL_KNOWN_BASES)
            is_direct_via = (not via_normalized or via_normalized == pol_normalized)

            should_add = False
            # Eğer bilinen bir ana limansa ve fiyat mantıklıysa (0 veya çok düşük değilse)
            if is_known_base and p20 > 500:
                # Via varsa ve Via farklı bir ana limansa, bu direct değildir.
                if not any(b == via_normalized for b in OOCL_KNOWN_BASES if b != pol_normalized):
                    should_add = True
            elif is_direct_via and is_high_price:
                should_add = True

            if should_add:
                # Aynı liman için en yüksek fiyatı baz al (Guarantees vs olabilir)
                if pol_normalized not in main_port_rates or p20 > main_port_rates[pol_normalized]["p20"]:
                    main_port_rates[pol_normalized] = {"p20": p20, "p40": r["p40"], "p40hc": r["p40hc"], "orig_name": pol}
        
        print(f"    ⚓ [OOCL] Referans Ana Limanlar: {len(main_port_rates)} adet", flush=True)

        # 2. ADIM: Rotaları Eşleştir
        temp_oocl_routes = [] 

        for r in raw_extracted_routes:
            pol = r.get("pol", "").upper()
            via_raw = r.get("via", "").upper()
            via_normalized = normalize_port_name(via_raw) 
            pol_normalized = normalize_port_name(pol)
            p20 = r["p20"]
            row_to_add = None

            # Senaryo A: Zaten Ana Liman (Kendisi)
            if pol_normalized in main_port_rates:
                # Fiyat kontrolü: Ana referans fiyatın %90'ından düşük değilse al
                if p20 >= (main_port_rates[pol_normalized]["p20"] * 0.9): 
                    row_to_add = r

            # Senaryo B: Feeder -> Main Bağlantısı
            elif via_normalized and via_normalized in main_port_rates:
                found_main = main_port_rates[via_normalized]
                
                # Feeder Fiyatı + Ana Liman Fiyatı
                total_20 = p20 + found_main["p20"]
                total_40 = r["p40"] + found_main["p40"]
                total_40hc = r["p40hc"] + found_main["p40hc"]
                
                new_r = r.copy()
                new_r["p20"] = total_20
                new_r["p40"] = total_40
                new_r["p40hc"] = total_40hc
                new_r["via"] = found_main["orig_name"] # Via artık ana liman ismi oldu
                row_to_add = new_r

            if row_to_add:
                if target_pods:
                    for t_pod in target_pods:
                        nr = row_to_add.copy()
                        nr["pod"] = t_pod
                        temp_oocl_routes.append(nr)
                else:
                    temp_oocl_routes.append(row_to_add)

        # 3. ADIM: Duplicate (Tekrar Eden) Temizliği
        unique_route_map = {} 
        for r in temp_oocl_routes:
            key = (r['pol'], r['via'], r['pod'])
            if key in unique_route_map:
                existing = unique_route_map[key]
                # Eğer aynı rotadan varsa, fiyatı yüksek olanı tut (Garanti olsun)
                if r['p40'] > existing['p40']: 
                    unique_route_map[key] = r
            else:
                unique_route_map[key] = r
        
        final_routes = list(unique_route_map.values())

    else:
        # -----------------------------------------------------------------
        # STANDART ROTA İŞLEME (Yang Ming, MSC, COSCO, SeaLead vb.)
        # -----------------------------------------------------------------
        for r in raw_extracted_routes:
             # Eğer POD (Varış Limanı) gemini tarafından bulunamadıysa 
             # ama mailin başında "target_pods" (İstanbul, Mersin vs) bulunduysa hepsine çoğalt.
             if not r["pod"] and target_pods:
                 for t_pod in target_pods:
                     nr = r.copy()
                     nr["pod"] = t_pod
                     final_routes.append(nr)
             else:
                 # POD varsa, standartlaştır
                 if r["pod"]: 
                     r["pod"] = global_map_pod(r["pod"], target_pods_context=target_pods)
                 final_routes.append(r)
    if not final_routes:
        final_routes = raw_extracted_routes # Basit fallback

    # =========================================================================
    # EXPORT DÖNGÜSÜ (SEALEAD DÜZELTMESİ İLE)
    # =========================================================================
    export_rows = []
    is_ym_msg = ("YANG" in filename) or ("YM" in filename) or ("YANG" in file_prefix) or GLOBAL_YANGMING_CONTEXT["is_active"]
    
    is_one_msg = "ONE" in filename or "ONE" in file_prefix
    is_zim_msg = "ZIM" in filename or "ZIM" in file_prefix
    is_sealead_msg = is_sealead_file
    is_msc_msg = "MSC" in filename or "MSC" in file_prefix
    is_cosco_msg = "COSCO" in filename or "COSCO" in file_prefix

    EUR_TO_USD = 1.00 

    for r in final_routes:
        p20 = safe_float(r.get("p20", 0))
        p40 = safe_float(r.get("p40", 0))
        p40hc = safe_float(r.get("p40hc", 0))
        
        # --- ONE LINE ---
        if is_one_msg and GLOBAL_ONE_CONTEXT["is_active"]:
            add = GLOBAL_ONE_CONTEXT["obs"] + GLOBAL_ONE_CONTEXT["pss"] + GLOBAL_ONE_CONTEXT["est"]
            p20 += add; p40 += (add + GLOBAL_ONE_CONTEXT["obs"]); p40hc += (add + GLOBAL_ONE_CONTEXT["obs"])
            
        # --- YANG MING (ETS DÜZELTMESİ) ---
        if is_ym_msg and GLOBAL_YANGMING_CONTEXT["is_active"]:
            # 1. Sabit Masraflar (ISPS + PSS) Herkese eklenir
            base_add = GLOBAL_YANGMING_CONTEXT["isps"] + GLOBAL_YANGMING_CONTEXT["pss"]
            p20 += base_add; p40 += base_add; p40hc += base_add
            
            # 2. ETS HESABI (TEU BAZLI)
            # ETS Birim Fiyatını (TEU Başına) Belirle
            ets_unit_eur = GLOBAL_YANGMING_CONTEXT["ets_20"]
            # Eğer 20'lik girilmemişse, 40'lık fiyatının yarısını al
            if ets_unit_eur == 0 and GLOBAL_YANGMING_CONTEXT["ets_40"] > 0:
                ets_unit_eur = GLOBAL_YANGMING_CONTEXT["ets_40"] / 2
            
            ets_teu_usd = ets_unit_eur * EUR_TO_USD
            
            # ÖNCE HERKESE EKLE
            p20 += ets_teu_usd
            p40 += (ets_teu_usd * 2)
            p40hc += (ets_teu_usd * 2)

            # 3. ÇİN HARİÇ KONTROLÜ (GERİ ÇIKARTMA)
            # Eğer "Çin Hariç" aktifse ve liman yasaklı listedeyse geri çıkart
            if GLOBAL_YANGMING_CONTEXT["ets_exclude_china"]:
                # Genişletilmiş Çin Limanları Listesi
                excluded_keywords = [
                    "XIAMEN", "FUZHOU", "CHANGZHOU", "CHONGQING", 
                    "XINGANG", "DALIAN", "QINGDAO", "LIANYUNGANG", 
                    "HUANGPU", "NANSHA", "NANTONG", "NANJING", "ZHANGJIAGANG", 
                    "SHANGHAI", "NINGBO", "SHEKOU", "SHANSHUI", "SHANTOU", 
                    "ZHONGSHAN", "XIAOLAN", "SANSHAN", "JIANGMEN", "FOSHAN", 
                    "JUIJIANG", "WUHU", "WUHAN", "WENZHOU", "YANGZHOU", 
                    "BEIJIAO", "LEILU", "RONGQI", "CIVET", "DOUMEN", 
                    "GAOLAN", "ZHAPU", "ZHANJIANG", "TIANJIN", "YANTIAN"
                ]
                
                current_pol_upper = str(r.get("pol", "")).upper()
                
                # Liman ismi listedeki kelimelerden birini içeriyor mu?
                is_in_exclusion = False
                for k in excluded_keywords:
                    if k in current_pol_upper:
                        is_in_exclusion = True
                        break
                
                if is_in_exclusion:
                    # ETS tutarını geri düşüyoruz (20'lik 1x, 40'lık 2x)
                    p20 -= ets_teu_usd
                    p40 -= (ets_teu_usd * 2)
                    p40hc -= (ets_teu_usd * 2)

        # --- ZIM ---
        if is_zim_msg and GLOBAL_ZIM_CONTEXT["is_active"]:
            added_val = GLOBAL_ZIM_CONTEXT["extra_total"]
            p20 += added_val; p40 += added_val; p40hc += added_val

        # --- SEA LEAD ---
        if is_sealead_msg and GLOBAL_SEALEAD_CONTEXT["is_active"]:
            if p40hc > 0: p40 = p40hc 
            eca_val = GLOBAL_SEALEAD_CONTEXT["eca"]
            p20 += eca_val; p40 += (eca_val * 2); p40hc += (eca_val * 2)

        # --- MSC ---
        if is_msc_msg and GLOBAL_MSC_CONTEXT["is_active"]:
            base_msc_add = GLOBAL_MSC_CONTEXT["cls"] + GLOBAL_MSC_CONTEXT["cdd"]
            msc_china_ports = ["SHANGHAI", "NINGBO", "QINGDAO", "XIAMEN", "SHEKOU", "YANTIAN", "NANSHA", "DALIAN", "XINGANG"]
            is_china_pol = any(cp in r["pol"].upper() for cp in msc_china_ports)
            crs_val = GLOBAL_MSC_CONTEXT["crs_china"] if is_china_pol else GLOBAL_MSC_CONTEXT["crs_other"]
            total_msc = base_msc_add + crs_val
            p20 += total_msc; p40 += total_msc; p40hc += total_msc

        # --- COSCO ---
        if is_cosco_msg and GLOBAL_COSCO_CONTEXT["is_active"]:
            ets_eur = GLOBAL_COSCO_CONTEXT["ets_eur"]
            if ets_eur > 0:
                ets_usd_20 = ets_eur * EUR_TO_USD
                ets_usd_40 = (ets_eur * 2) * EUR_TO_USD
                p20 += ets_usd_20; p40 += ets_usd_40; p40hc += ets_usd_40

        export_rows.append({
            "POL": r["pol"], "VIA": r.get("via", ""), "POD": r.get("pod", ""), "TOCITY": "", "CURR": r.get("curr", "USD"),
            "FREETIME": "16", "TRANSIT": "", "CUSTOMERDESCRIPTION": "",
            "TYPE_1": "20 DC", "AMOUNT_1": round(p20, 2),
            "TYPE_2": "40 DC", "AMOUNT_2": round(p40, 2),
            "TYPE_3": "40 HC", "AMOUNT_3": round(p40hc, 2)
        })

    return export_rows

# =============================================================================
# EXCEL İŞLEME
# =============================================================================

def process_excel_file(file_path):
    print(f"  > Excel İşleniyor: {os.path.basename(file_path)}", flush=True)
    fname = os.path.basename(file_path).upper()
    is_one_file = "ONE" in fname
    is_ym_file = "YANG" in fname
    is_msc_file = "MSC" in fname
    
    # --- ONE LINE / YANG MING LOGIC ---
    one_add20 = 0; one_add40 = 0
    if is_one_file and GLOBAL_ONE_CONTEXT["is_active"]:
        one_add20 = GLOBAL_ONE_CONTEXT["obs"] + GLOBAL_ONE_CONTEXT["pss"] + GLOBAL_ONE_CONTEXT["est"]
        one_add40 = (GLOBAL_ONE_CONTEXT["obs"] * 2) + GLOBAL_ONE_CONTEXT["pss"] + GLOBAL_ONE_CONTEXT["est"]

    ym_add20 = 0; ym_add40 = 0; ym_exclude_china = False
    if is_ym_file and GLOBAL_YANGMING_CONTEXT["is_active"]:
        base_add = GLOBAL_YANGMING_CONTEXT["isps"] + GLOBAL_YANGMING_CONTEXT["pss"]
        ets_20_usd = GLOBAL_YANGMING_CONTEXT["ets_20"] * EUR_TO_USD
        ets_40_usd = GLOBAL_YANGMING_CONTEXT["ets_40"] * EUR_TO_USD
        ym_add20 = base_add + ets_20_usd
        ym_add40 = base_add + ets_40_usd
        ym_exclude_china = GLOBAL_YANGMING_CONTEXT["ets_exclude_china"]

    try:
        xl = pd.ExcelFile(file_path)
        all_rows = []
        for sheet in xl.sheet_names:
            if any(bad in sheet.upper() for bad in EXCEL_SHEET_BLACKLIST): continue
            
            # ================================================================
            # MSC ÖZEL EXCEL OKUMA MANTIĞI
            # ================================================================
            if is_msc_file:
                print(f"    🔍 [MSC Analiz] Sayfa taranıyor: {sheet}", flush=True)
                df_raw_scan = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=30)
                found_pod_in_excel = None
                for i, row in df_raw_scan.iterrows():
                    row_text = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                    for tp in TURKISH_PORTS:
                        if tp in row_text:
                            found_pod_in_excel = global_map_pod(tp); break
                    if found_pod_in_excel: break
                if not found_pod_in_excel: found_pod_in_excel = "ISTANBUL"

                header_idx = None
                for i, row in df_raw_scan.iterrows():
                    row_str = " ".join([str(x).upper() for x in row.values])
                    if ("20" in row_str or "TEU" in row_str) and ("40" in row_str or "FEU" in row_str):
                        header_idx = i; break
                
                if header_idx is None: 
                    print(f"    ⚠️ [MSC HATA] Fiyat başlık satırı bulunamadı, bu sayfa atlanıyor.")
                    continue

                df = pd.read_excel(file_path, sheet_name=sheet, header=header_idx)
                df.columns = [str(c).upper().strip() for c in df.columns]
                col_map = {}
                for col in df.columns:
                    c = col.replace("'", "").replace(" ", "").replace(".", "")
                    if "NOR" in c: continue 
                    if "20" in c and ("DV" in c or "DRY" in c or "GP" in c or c=="20" or "TEU" in c): col_map[col] = "20DV"
                    elif "40" in c and "HC" in c: col_map[col] = "40HC"
                    elif "40" in c and ("DV" in c or "DRY" in c or "GP" in c or c=="40" or "FEU" in c): col_map[col] = "40DV"

                pol_col = None
                candidate_cols = [c for c in df.columns if c not in col_map]
                for col in candidate_cols:
                    sample_values = df[col].dropna().head(10).astype(str).tolist()
                    score_valid = 0; score_invalid = 0
                    for val in sample_values:
                        val = val.strip().upper()
                        if val.replace('.', '', 1).isdigit(): score_invalid += 10
                        elif val in ["USD", "EUR", "VALID", "TOTAL"]: score_invalid += 5
                        elif len(val) < 3: score_invalid += 1
                        elif is_turkish_port(val): score_invalid += 10
                        else: score_valid += 1
                    if score_valid > 0 and score_invalid <= 2:
                        pol_col = col; break 
                
                if not pol_col: continue
                print(f"    📊 [MSC DEBUG] Sütunlar -> POL: {pol_col} | POD: {found_pod_in_excel}", flush=True)

                for idx, row in df.iterrows():
                    pol_raw = str(row[pol_col]).strip().upper()
                    if pol_raw == "NAN" or pol_raw == "NONE" or len(pol_raw) < 3: continue
                    if "PRICE" in pol_raw or "VALID" in pol_raw: continue
                    
                    pol_clean = global_clean_port_name(pol_raw)
                    if is_turkish_port(pol_clean): continue

                    def get_price(k):
                        cols = [c for c, v in col_map.items() if v == k]
                        if not cols: return 0
                        val = row[cols[0]]
                        try:
                            val_clean = re.sub(r'[^\d.,]', '', str(val))
                            if not val_clean: return 0
                            return float(val_clean)
                        except: return 0

                    p20 = get_price("20DV"); p40 = get_price("40DV"); p40hc = get_price("40HC")
                    if p20 == 0 and p40 == 0: continue

                    is_china = any(cp in pol_clean for cp in CHINA_PORTS)
                    cls_val = GLOBAL_MSC_CONTEXT["cls"]             
                    crs_china_val = GLOBAL_MSC_CONTEXT["crs_china"] 
                    crs_other_val = GLOBAL_MSC_CONTEXT["crs_other"] 
                    cdd_val = GLOBAL_MSC_CONTEXT["cdd"]             

                    per_teu_extra = 0
                    if is_china: per_teu_extra = cls_val + crs_china_val
                    else: per_teu_extra = crs_other_val

                    add_20 = per_teu_extra * 1
                    add_40 = per_teu_extra * 2
                    
                    final_20 = p20 + add_20 + cdd_val
                    final_40 = p40 + add_40 + cdd_val
                    final_40hc = p40hc + add_40 + cdd_val

                    all_rows.append({
                        "POL": pol_clean, "VIA": "", "POD": found_pod_in_excel, "TOCITY": "", "CURR": "USD",
                        "FREETIME": "", "TRANSIT": "", "CUSTOMERDESCRIPTION": "MSC",
                        "TYPE_1": "20 DC", "AMOUNT_1": final_20,
                        "TYPE_2": "40 DC", "AMOUNT_2": final_40,
                        "TYPE_3": "40 HC", "AMOUNT_3": final_40hc
                    })
                continue 

            # ================================================================
            # STANDART (ESKİ) EXCEL OKUMA MANTIĞI
            # ================================================================
            df_temp = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=25)
            header_idx = -1
            for i, row in df_temp.iterrows():
                row_str = " ".join([str(x).upper() for x in row.values])
                if (("POL" in row_str or "LOAD" in row_str or "POR" in row_str) and ("20" in row_str or "PRICE" in row_str)):
                    header_idx = i; break
            if header_idx == -1: continue
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_idx)
            df.columns = [str(c).upper().strip() for c in df.columns]
            pol_col = next((c for c in df.columns if "POR DESCRIPTION" in c), None)
            if not pol_col: pol_col = next((c for c in df.columns if "POL" in c or "LOAD" in c or "ORIGIN" in c), None)
            pod_col = next((c for c in df.columns if "DEL DESCRIPTION" in c or "DELIVERY DESCRIPTION" in c), None)
            if not pod_col: pod_col = next((c for c in df.columns if "POD" in c or "DISCH" in c or "DEST" in c), None)
            col_commodity = next((c for c in df.columns if "COMMODITY GROUP NAME" in c), None)
            via_col = next((c for c in df.columns if "VIA" in c), None)
            col_20 = next((c for c in df.columns if "20" in c and "PRICE" not in c and "QTY" not in c), None)
            col_40 = next((c for c in df.columns if "40" in c and "HC" not in c and "PRICE" not in c), None)
            col_40hc = next((c for c in df.columns if ("HC" in c or "HQ" in c) and "PRICE" not in c), None)
            if not (pol_col and (col_20 or col_40)): continue
            
            for _, row in df.iterrows():
                if pd.isna(row[pol_col]): continue
                via_val = ""
                try:
                    p20 = safe_float(row[col_20]) if col_20 and pd.notna(row[col_20]) else 0
                    p40 = safe_float(row[col_40]) if col_40 and pd.notna(row[col_40]) else 0
                    p40hc = safe_float(row[col_40hc]) if col_40hc and pd.notna(row[col_40hc]) else 0
                except: continue
                if p20 < 1000 and p40 < 1000: continue
                
                raw_pol_str = str(row[pol_col])
                pol_parts = [p.strip() for p in raw_pol_str.split('/') if len(p.strip()) > 1]
                
                for single_pol in pol_parts:
                    pol_clean = global_clean_port_name(single_pol)
                    if len(pol_clean) < 2: continue
                    if is_turkish_port(pol_clean): continue

                    raw_pod_cell = str(row[pod_col]).strip() if pod_col and pd.notna(row[pod_col]) else "ISTANBUL"
                    final_pod = global_map_pod(raw_pod_cell)
                   # ============================================================
                    # SADECE ONE LINE İÇİN ÖZEL MANTIK
                    # ============================================================
                    via_val = ""

                    if is_one_file:
    # -------------------------------------------------------------------------
    # 1. VIA KONTROLÜ 
    # -------------------------------------------------------------------------
                        if via_col and pd.notna(row[via_col]):
                            raw_via = str(row[via_col]).strip().upper()
                            if raw_via not in ["NAN", "NONE", "0", ""]:
                                via_val = PORT_CODE_MAP.get(raw_via, raw_via)

    # -------------------------------------------------------------------------
    # 2. POD DÜZELTMESİ (ISTANBUL -> MARPORT/HAYDARPAŞA)
    # -------------------------------------------------------------------------
    # One Line dosyalarında POD sütunu boş olabilir, bu yüzden "DEL Description"
    # sütununa da bakmamız gerekiyor.
    
    # 'DEL Description' sütunundaki veriyi güvenli bir şekilde al:
    # 
                        del_desc_val = str(row.get('DEL Description', '')).strip().upper()
                        current_pod_val = str(raw_pod_cell).strip().upper()

    # Kontrol: Hem POD hücresinde hem de DEL Description sütununda "ISTANBUL" arıyoruz.
                        if col_commodity and (("ISTANBUL" in current_pod_val) or ("ISTANBUL" in del_desc_val)):
        
        # Commodity Group Name sütunundaki değeri al (Örn: MARPORT FAK)
                            comm_val = str(row[col_commodity]).strip()
        
                            if comm_val and comm_val.upper() != "NAN":
            # Değeri olduğu gibi POD'a ata
                                final_pod = comm_val
            
                                final_pod = comm_val.replace(" FAK", "").replace("fak", "").strip()
                    # ============================================================
                    
                    final_20 = p20; final_40 = p40; final_40hc = p40hc
                    
                    if is_one_file:
                        final_20 += one_add20; final_40 += one_add40; final_40hc += one_add40
                    
                    if is_ym_file:
                        is_china = any(cp in pol_clean.upper() for cp in CHINA_PORTS)
                        current_add20 = ym_add20
                        current_add40 = ym_add40
                        if ym_exclude_china and is_china:
                            ets_20_usd = GLOBAL_YANGMING_CONTEXT["ets_20"] * EUR_TO_USD
                            ets_40_usd = GLOBAL_YANGMING_CONTEXT["ets_40"] * EUR_TO_USD
                            current_add20 -= ets_20_usd
                            current_add40 -= ets_40_usd

                        final_20 += current_add20; final_40 += current_add40; final_40hc += current_add40

                    all_rows.append({
                        "POL": pol_clean, "VIA": via_val, "POD": final_pod, "TOCITY": "", "CURR": "USD",
                        "FREETIME": "", "TRANSIT": "", "CUSTOMERDESCRIPTION": "",
                        "TYPE_1": "20 DC", "AMOUNT_1": final_20,
                        "TYPE_2": "40 DC", "AMOUNT_2": final_40,
                        "TYPE_3": "40 HC", "AMOUNT_3": final_40hc
                    })
        return all_rows
    except Exception as e:
        print(f"Hata: {e}")
        return None

# =============================================================================
# CSV KAYIT
# =============================================================================

def save_to_csv(data, filename):
    if not data: return

    df = pd.DataFrame(data)


   # # ---------------------------------------------------------
    # ZIM ENTEGRASYONU 
    # ---------------------------------------------------------
    if "ZIM" in filename.upper():
        print("    ⚙️ [ZIM] V10: 'DIRECT' yazıları temizleniyor, hücreler boş bırakılıyor...", flush=True)

        # -----------------------------------------------------
        # 1. AYARLAR & HARİTALAR
        # -----------------------------------------------------
        
        # MANTIKLI FİYAT EŞİĞİ 
        MIN_TOTAL_FREIGHT_THRESHOLD = 1000.0 

        via_code_to_main_port = {
            "KRPUS": "BUSAN", "PUSAN": "BUSAN", "BUSAN": "BUSAN",
            "CNSNH": "SHANGHAI", "SHANGHAI": "SHANGHAI",
            "CNNGB": "NINGBO", "CNSGB": "NINGBO", "NINGBO": "NINGBO",
            "CNQIN": "QINGDAO", "QINGDAO": "QINGDAO",
            "CNSAD": "DA CHAN BAY", "DA CHAN BAY": "DA CHAN BAY",
            "VNTCT": "CAI MEP", "VNHCM": "CAI MEP", "CAI MEP": "CAI MEP",
            "THLEM": "LAEM CHABANG",
            "CNTXG": "XINGANG", "CNYTN": "YANTIAN",
            "HKHKG": "HONG KONG", "SGSIN": "SINGAPORE"
        }

        zim_pol_map = {
            "CNXNG": "XINGANG", "CNXING": "XINGANG", "CNTXG": "XINGANG", 
            "CNDAL": "DALIAN", "CNXIA": "XIAMEN",
            "TWKLG": "KEELUNG", "TWKSG": "KAOHSIUNG", "TWTXG": "TAICHUNG", "TWTYG": "TAOYUAN",
            "JPTYO": "TOKYO", "JPNGO": "NAGOYA", "JPYOK": "YOKOHAMA", "JPKBE": "KOBE", "JPOSA": "OSAKA",
            "SGSIN": "SINGAPORE", "IDJKT": "JAKARTA", "IDSUB": "SURABAYA", 
            "VNHAI": "HAIPHONG", "VNDAD": "DA NANG", 
            "THLEM": "LAEM CHABANG", "THBKK": "BANGKOK", "THLKR": "LAT KRABANG",
            "MYPKL": "PORT KLANG", "PHZMP": "MANILA", "PHMNL": "MANILA", 
            "KHPKH": "PHNOM PENH", "KHOUX": "SIHANOUKVILLE", "HKHKG": "HONG KONG",
            "CNQNQ": "QINGDAO", "CNOHG": "HUANGPU", "CNCHZ": "CHANGZHOU", "CNOCQ": "CHONGQING",
            "CNFUZ": "FUZHOU", "CNJIC": "JIAOXIN", "CNJJG": "JIUJIANG", "CNLYU": "LIANYUNGANG",
            "CNNHJ": "NANHAI", "CNNNJ": "NANJING", "CNNTG": "NANTONG", "CNZJG": "ZHANGJIAGANG",
            "CNZAP": "ZHAPU", "CNCHE": "CHANGSHA", "CNSHS": "SHANTOU", "CNTAG": "TAICANG",
            "CNTZO": "TAIZHOU", "CNWZU": "WENZHOU", "CNWUH": "WUHAN", "CNWUU": "WUHU",
            "CNYAZ": "YANGZHOU", "CNYIH": "YICHANG", "CNYEY": "YUEYANG", "CNZEN": "ZHENJIANG",
            "CNWAZ": "WANZHOU", "CNLUZ": "LUZHOU", "CNYND": "YANTAI", "CNDAF": "DA FENG",
            "CNQIZ": "QINZHOU", "CNGAO": "GAOLAN", "CNJJA": "JIAOJIANG", "CNSHU": "SHUNDE",
            "CNIHZ": "HUANGSHI", "CNZHQ": "ZHAOQING", "CNLEL": "LIULIN", "CNHAU": "HUADU",
            "CNHUA": "HUADU", "CNJIA": "JIANGMEN", "CNNSJ": "NANSHA", "CNRQI": "RONGQI",
            "CNBJO": "BEIJIAO", "CNSHN": "SHANTOU", "CNXIO": "XIAOLAN", "CNZHJ": "ZHANJIANG",
            "CNZSH": "ZHONGSHAN", "CNZHU": "ZHUHAI", "CNOZX": "OUZHU", "CNWUZ": "WUZHOU",
            "CNHKG": "HONG KONG"
        }

        def normalize_pod(pod_name):
            p = pod_name.upper().strip()
            if "AMBARLI" in p: return "AMBARLI"
            if "DERINCE" in p: return "DERINCE"
            if "IZMIT" in p: return "DERINCE"
            if "ISTANBUL" in p or "TRIST" in p: return "ISTANBUL"
            return p

        base_prices_map = {}
        final_valid_rows = []

        # -----------------------------------------------------
        # 2. AŞAMA: BASE FİYATLARI TOPLA
        # -----------------------------------------------------
        print("    🔎 [ZIM] Ana Liman (Base Port) fiyatları hafızaya alınıyor...")
        
        for index, row in df.iterrows():
            pol_raw = str(row.get('POL', '')).strip().upper()
            via_raw = str(row.get('VIA', '')).strip().upper()
            
            # Base Navlun Tespiti
            is_base_structure = (via_raw in ["", "NAN", "NONE", "-", "DIRECT"] or via_raw == pol_raw)
            
            prices = {'20': 0, '40': 0, '40HC': 0}
            max_p = 0
            for i in range(1, 6):
                t_col = f"TYPE_{i}"; a_col = f"AMOUNT_{i}"
                if t_col in df.columns and a_col in df.columns:
                    t_val = str(row[t_col]).upper()
                    a_val = safe_float(row[a_col])
                    if a_val > max_p: max_p = a_val
                    if "20" in t_val: prices['20'] = a_val
                    elif "40" in t_val and "HC" in t_val: prices['40HC'] = a_val
                    elif "40" in t_val: prices['40'] = a_val
            
            # Emniyet: Fiyat çok yüksekse ve yapısal olarak karışık değilse Base kabul et
            if max_p > 2000: is_base_structure = True

            if is_base_structure and max_p > 100:
                pols = [p.strip() for p in pol_raw.replace('>', '/').split('/') if p.strip()]
                raw_pods = [p.strip() for p in str(row.get('POD', '')).replace('>', '/').split('/') if p.strip()]

                for r_pol in pols:
                    clean_pol = zim_pol_map.get(r_pol, r_pol)
                    if clean_pol not in base_prices_map: base_prices_map[clean_pol] = {}
                    
                    for r_pod in raw_pods:
                        std_pod = normalize_pod(r_pod)
                        targets = [std_pod]
                        if std_pod == "ISTANBUL": targets = ["AMBARLI", "DERINCE"]
                        
                        for t in targets:
                            if t not in base_prices_map[clean_pol] or prices['40HC'] > base_prices_map[clean_pol][t]['40HC']:
                                base_prices_map[clean_pol][t] = prices.copy()

        # -----------------------------------------------------
        # 3. AŞAMA: SATIRLARI OLUŞTUR (Add-on Hesapla & Direct Temizle)
        # -----------------------------------------------------
        print(f"    ℹ️  [ZIM] {len(base_prices_map)} Ana Liman hafızada. Tablo oluşturuluyor...")
        
        for index, row in df.iterrows():
            raw_via = str(row.get('VIA', '')).strip().upper()
            raw_pol = str(row.get('POL', '')).strip().upper()
            raw_pod = str(row.get('POD', '')).strip().upper()

            pols = [p.strip() for p in raw_pol.replace('>', '/').split('/') if p.strip()]
            vias = [v.strip() for v in raw_via.replace('>', '/').split('/') if v.strip()]
            
            # Eğer VIA boşsa işlem yapmak için geçici olarak 'DIRECT' diyoruz
            # AMA çıktıya yazarken bunu boşaltacağız.
            if not vias: vias = ["DIRECT"] 
            
            pods = [d.strip() for d in raw_pod.replace('>', '/').split('/') if d.strip()]

            # Satırın ham fiyatlarını oku
            addon_prices = {'20': 0, '40': 0, '40HC': 0}
            row_max_val = 0
            for i in range(1, 6):
                t_col = f"TYPE_{i}"; a_col = f"AMOUNT_{i}"
                if t_col in df.columns and a_col in df.columns:
                    t_val = str(row[t_col]).upper()
                    a_val = safe_float(row[a_col])
                    if a_val > row_max_val: row_max_val = a_val

            for current_pol in pols:
                clean_pol = zim_pol_map.get(current_pol, current_pol)
                
                for current_via in vias:
                    
                    target_hub = via_code_to_main_port.get(current_via)
                    
                    # --- SENARYO 1: ADD-ON  ---
                    if target_hub and target_hub in base_prices_map:
                        base_hub_prices = base_prices_map[target_hub]
                        
                        destinations = []
                        if any("ISTANBUL" in normalize_pod(p) for p in pods):
                            destinations = list(base_hub_prices.keys())
                        else:
                            for p in pods:
                                std = normalize_pod(p)
                                if std in base_hub_prices: destinations.append(std)
                        
                        for dest_pod in destinations:
                            base_price = base_hub_prices[dest_pod]
                            new_row = row.copy()
                            is_valid = False
                            
                            for i in range(1, 6):
                                t_col = f"TYPE_{i}"; a_col = f"AMOUNT_{i}"
                                if t_col in df.columns and a_col in df.columns:
                                    t_val = str(row[t_col]).upper()
                                    raw_addon = safe_float(row[a_col])
                                    
                                    total_val = 0
                                    if "20" in t_val: total_val = base_price['20'] + raw_addon
                                    elif "40" in t_val and "HC" in t_val: total_val = base_price['40HC'] + raw_addon
                                    elif "40" in t_val: total_val = base_price['40'] + raw_addon
                                    
                                    if total_val > MIN_TOTAL_FREIGHT_THRESHOLD:
                                        new_row[a_col] = total_val
                                        is_valid = True
                                    else:
                                        new_row[a_col] = 0
                            
                            if is_valid:
                                new_row['POL'] = clean_pol
                                new_row['VIA'] = target_hub
                                new_row['POD'] = dest_pod
                                final_valid_rows.append(new_row)

                    # --- SENARYO 2: DIRECT / BASE SATIR ---
                    # --- VIA=DIRECT ise veya Boş ise    ---
                    elif current_via == "DIRECT" or current_via == "" or (current_via == current_pol):
                        if row_max_val > MIN_TOTAL_FREIGHT_THRESHOLD:
                            for p in pods:
                                std_pod = normalize_pod(p)
                                targets = [std_pod]
                                if std_pod == "ISTANBUL": targets = ["AMBARLI", "DERINCE"]
                                for t in targets:
                                    new_row = row.copy()
                                    new_row['POL'] = clean_pol
                                    
                                  
                                    new_row['VIA'] = "" 
                                    
                                    new_row['POD'] = t
                                    final_valid_rows.append(new_row)

        if final_valid_rows:
            df = pd.DataFrame(final_valid_rows)
            cols_to_check = ['POL', 'POD', 'VIA', 'AMOUNT_1']
            cols_existing = [c for c in cols_to_check if c in df.columns]
            df = df.drop_duplicates(subset=cols_existing)
            
            print(f"    ✅ [ZIM] V10 Tamamlandı. 'Direct' yazıları temizlendi. Satır: {len(df)}")

            
    cols = [c for c in BASE_HEADER_COLS if c in df.columns]
    type_amount_cols = []
    for i in range(1, 20):
        t, a = f"TYPE_{i}", f"AMOUNT_{i}"
        if t in df.columns and a in df.columns:
            type_amount_cols.extend([t, a])
    
    final_cols = cols + type_amount_cols
    df = df.reindex(columns=final_cols)
    
    header_str = ";".join(final_cols)
    
    try:
        with open(filename, "w", encoding="utf-8-sig", newline="") as f:
            f.write(header_str + "\n")
            df.to_csv(f, sep=";", index=False, header=False)
        print(f"  ✅ OLUŞTURULDU: {filename}", flush=True)
    except Exception as e:
        print(f"  ❌ HATA (CSV Kayıt): {e}", flush=True)


def clean_msc_duplicates_only():
    """
    Klasördeki CSV dosyalarını tarar.
    SADECE dosya isminde 'MSC' geçen dosyalara işlem yapar.
    Bu dosyalarda POL, Via, POD (ilk 3 sütun) aynı ise;
    - İlk satırı tutar.
    - Diğerlerini siler.
    """
    print("-" * 60)
    print("🧹 MSC ÖZEL KONTROL: Sadece MSC dosyalarındaki kopyalar temizleniyor...")

    # Klasördeki tüm csv dosyalarını bul
    csv_files = glob.glob("*.csv")

    for f in csv_files:
        
        if "MSC" not in f.upper():
            continue

        try:
            
            df = pd.read_csv(f, sep=None, engine='python')

            # En az 3 sütun var mı kontrol et
            if len(df.columns) < 3:
                continue

            # İlk 3 sütunun başlığını al (POL, VIA, POD varsayımıyla)
            first_three_cols = df.columns[:3].tolist()
            
            initial_count = len(df)

            # --- SADECE MSC İÇİN TEMİZLİK ---
            df.drop_duplicates(subset=first_three_cols, keep='first', inplace=True)

            final_count = len(df)

            
            if initial_count > final_count:
                diff = initial_count - final_count
                df.to_csv(f, index=False, sep=';', encoding='utf-8-sig')
                print(f"   ✅ {f} (MSC): {diff} adet kopya satır temizlendi.")
            else:
                print(f"   ℹ️ {f} (MSC): Temizlenecek kopya satır bulunamadı.")

        except Exception as e:
            print(f"   ⚠️ Hata ({f}): {e}")

    print("-" * 60)


def clear_via_column_in_all_outputs():
    """
    Çalışma dizinindeki tüm CSV dosyalarını gezer ve 
    'VIA' sütunundaki değerleri siler.
    """
    import os
    import pandas as pd
    import glob

    # Mevcut dizindeki tüm CSV dosyalarını bul
    csv_files = glob.glob("*.csv")
    
    if not csv_files:
        print("\n🔍 Temizlenecek CSV dosyası bulunamadı.")
        return

    print(f"\n🧹 VIA sütunu temizleme işlemi başlatıldı ({len(csv_files)} dosya)...")

    for file in csv_files:
        try:
            # Dosyayı oku (Ayırıcıyı tespit etmeye çalış veya yaygın kullanılan ';' ve ',' dene)
            # Kodunda genellikle ';' kullanıldığı için önce onu deniyoruz
            try:
                df = pd.read_csv(file, sep=';', encoding='utf-8-sig')
                # Eğer POL sütunu yoksa yanlış ayırıcı olabilir, virgül dene
                if 'POL' not in df.columns and len(df.columns) < 2:
                    df = pd.read_csv(file, sep=',', encoding='utf-8-sig')
            except:
                df = pd.read_csv(file, sep=',', encoding='utf-8-sig')

            # Sütun isimlerini büyük harfe çevirerek 'VIA'yı bul
            df.columns = [str(c).strip().upper() for c in df.columns]

            if 'VIA' in df.columns:
                # VIA sütunundaki tüm değerleri boş string yap
                df['VIA'] = ""
                
                # Dosyayı orijinal formatında (noktalı virgül ile) tekrar kaydet
                df.to_csv(file, sep=';', index=False, encoding='utf-8-sig')
                print(f"   ✅ Temizlendi: {file}")
            else:
                print(f"   ℹ️  Atlandı (VIA sütunu yok): {file}")

        except Exception as e:
            print(f"   ❌ Hata oluştu ({file}): {e}")

    print("✨ Bütün dosyalar güncellendi.")

def ana_islem_dongusu():
    if 'cma_special_logic_processor' in globals():
        cma_special_logic_processor()


    mail_files = [f for f in glob.glob("*.*") if f.lower().endswith(('.msg', '.eml'))]
    print(f"✉️  MAİLLER İŞLENİYOR ({len(mail_files)} dosya)...", flush=True)

    for f in mail_files:
        if f.endswith(".py") or f.endswith(".csv"): continue
        
        filename_lower = os.path.basename(f).lower()

       
        if "cma" in filename_lower: continue

        # =========================================================
        #2. ZIM DOSYALARINI YAKALA VE ÖZEL İŞLE
        # =========================================================
        if "zim" in filename_lower:
            print(f"🚢 ZIM Dosyası Tespit Edildi: {f}")
            
            
            zim_data = zim_special_logic_processor(f)
            
            if zim_data:
                if 'apply_final_corrections' in globals():
                    zim_data = apply_final_corrections(zim_data, f)

                df_zim = pd.DataFrame(zim_data)

                cols = ["POL", "VIA", "POD", "TOCITY", "CURR", "FREETIME", "TRANSIT", 
                        "CUSTOMERDESCRIPTION", "TYPE_1", "AMOUNT_1", "TYPE_2", "AMOUNT_2", 
                        "TYPE_3", "AMOUNT_3"]
                
                for c in cols:
                    if c not in df_zim.columns: df_zim[c] = ""
                
                # CSV Olarak Kaydet
                base_name = os.path.splitext(f)[0]
                output_csv = f"{base_name}.csv"
                df_zim[cols].to_csv(output_csv, sep=';', index=False, encoding='utf-8-sig')
                
                print(f"✅ ZIM Çıktısı Hazır: {output_csv}")
            else:
                print(f"❌ ZIM verisi çıkarılamadı (Dosya: {f}). Manuel kontrol gerekebilir.")

            
            continue 
        
        # 3. Diğer firmalar (MSC, ONE, vb.) için standart işlem
        base_name = os.path.splitext(f)[0]
        output_csv = f"{base_name}.csv"
        
        # Genel işlemciyi çağır
        data = process_msg_file(f)
        
        if data: 
            # MSC ve ONE temizliği burada devreye giriyor
            if 'apply_final_corrections' in globals():
                data = apply_final_corrections(data, f)
            
            save_to_csv(data, output_csv)
    
    print("-" * 60)
    
   
    all_files_now = glob.glob("*.*")
    doc_files = [f for f in all_files_now if f.lower().endswith(('.xlsx', '.xls', '.pdf', '.eml'))]
    print(f"📂 DÖKÜMANLAR İŞLENİYOR ({len(doc_files)} dosya)...", flush=True)
    
    for f in doc_files:
        filename = os.path.basename(f)
        
       
        if f.endswith(".py") or f.endswith(".csv"): continue
        if "cma" in filename.lower(): continue

        # --- HAPAG KONTROLÜ  ---
        
        if "hapag" in filename.lower() and f.lower().endswith(".eml"):
            print(f"   ⚓ Hapag Modülü Çalışıyor: {filename}")
            process_hapag_special(f)
            continue  
        if "hapag" in filename and "add-on" in filename:
            print(f"   ℹ️  Bilgi: {filename} bir yardımcı dosyadır, tek başına işlenmez.")
            continue
        
        if f.lower().endswith(".eml"):
            continue

        
        base_name = os.path.splitext(f)[0]
        output_csv = f"{base_name}.csv"
        ext = os.path.splitext(f)[1].lower()
        data = None
        
        if ext == ".pdf": data = process_pdf_file(f)
        elif ext in [".xlsx", ".xls"]: data = process_excel_file(f)
        
        if data: 
            # Dökümanlar (PDF/Excel) için de temizlik çalışsın
            if 'apply_final_corrections' in globals():
                data = apply_final_corrections(data, f)
            save_to_csv(data, output_csv)
    
    # Varsa MSC temizliği
    if 'clean_msc_duplicates_only' in globals():
        clean_msc_duplicates_only()


    clear_via_column_in_all_outputs()
    apply_tmaxx_corrections_to_all_csvs("tmaxxliste.xlsx")
    final_pol_correction()

def main():
    st.title("⚓ Denizcilik Tarife İşleme Uygulaması")
    st.write("İşlenecek dosyaları (Excel, PDF, EML, MSG) ve varsa yardımcı dosyaları (Add-on listeleri) buraya sürükleyip bırakın.")

    # 1. STREAMLIT HAFIZASI (Session State) OLUŞTURMA
    # Sayfa yenilense bile dosyaların kaybolmamasını sağlar
    if 'uretilen_dosyalar' not in st.session_state:
        st.session_state['uretilen_dosyalar'] = []

    uploaded_files = st.file_uploader("Dosyaları Buraya Yükleyin", accept_multiple_files=True)

    if uploaded_files:
        if st.button("🚀 İşlemi Başlat"):
            with tempfile.TemporaryDirectory() as temp_dir:
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                original_cwd = os.getcwd()
                os.chdir(temp_dir)

                try:
                    with st.spinner("Dosyalar işleniyor..."):
                        ana_islem_dongusu() # Sizin kodunuz çalışıyor
                    
                    csv_files = glob.glob("*.csv")
                    
                    if csv_files:
                        st.success(f"✅ İşlem tamamlandı! {len(csv_files)} adet çıktı dosyası üretildi.")
                        
                        # 2. DOSYALARI HAFIZAYA KAYDETME
                        gecici_hafiza = []
                        for csv_file in csv_files:
                            with open(csv_file, "rb") as f:
                                gecici_hafiza.append({
                                    "isim": csv_file,
                                    "veri": f.read() # Dosyayı bayt olarak okuyup hafızaya alıyoruz
                                })
                        
                        st.session_state['uretilen_dosyalar'] = gecici_hafiza
                    else:
                        st.warning("⚠️ İşlem tamamlandı ancak çıktı oluşmadı.")
                
                except Exception as e:
                    st.error(f"❌ İşlem sırasında bir hata oluştu: {str(e)}")
                
                finally:
                    os.chdir(original_cwd)

    # 3. İNDİRME BUTONLARINI GÖSTERME (İşlemi Başlat butonunun DIŞINA alındı)
    if st.session_state['uretilen_dosyalar']:
        st.write("---")
        st.write("### 📂 Üretilen Dosyalar")
        
        # BONUS: Tüm dosyaları tek bir ZIP olarak indirme butonu
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for dosya in st.session_state['uretilen_dosyalar']:
                zf.writestr(dosya["isim"], dosya["veri"])
        
        st.download_button(
            label="📦 Tüm Dosyaları Tek Seferde İndir (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="islenmis_tarifeler.zip",
            mime="application/zip",
            type="primary" # Rengini belirgin yapar
        )
        
        st.write("**Veya tek tek indirebilirsiniz:**")
        
        # Dosyaları tek tek indirme butonları
        for dosya in st.session_state['uretilen_dosyalar']:
            st.download_button(
                label=f"📥 İndir: {dosya['isim']}",
                data=dosya["veri"],
                file_name=dosya["isim"],
                mime="text/csv",
                key=dosya["isim"] # Sayfa yenilendiğinde butonların karışmasını önler
            )

if __name__ == "__main__":
    main()
    






