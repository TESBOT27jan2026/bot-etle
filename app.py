import email
from email import encoders
from email.header import decode_header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import os
import smtplib
import time
from urllib.parse import urlencode

from bs4 import BeautifulSoup
import pandas as pd
import requests

# --- KONFIGURASI ---
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")

# Header HTTP agar request menyerupai browser resmi
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

# --- HEADER ASLI (Lengkap 24 Kolom) ---
FULL_HEADERS = [
    "No. Baris",
    "Plat Nomor",
    "No Rangka",
    "No Mesin",
    "Status ETLE",
    "Link Pengecekan",
    "Lokasi",
    "Tanggal",
    "Jenis Pelanggaran",
    "Status",
    "Link Detail",
    "Path Foto",
    "Detail Jenis Pelanggaran",
    "Dasar Hukum",
    "Waktu Kejadian",
    "Merk",
    "Tipe",
    "Warna",
    "Nomor Rangka Detail",
    "Model",
    "Tahun",
    "Nomor Mesin Detail",
    "STNK Berlaku Sampai",
    "Estimasi Denda",
]


# --- FUNGSI HITUNG DENDA (Sesuai Logika Asli) ---
def hitung_denda(jenis_pelanggaran):
  text = str(jenis_pelanggaran).lower()
  if any(x in text for x in ["handphone", "ponsel", "wajar"]):
    return 750000
  elif "helm" in text:
    return 250000
  elif any(x in text for x in ["sabuk", "safety belt"]):
    return 250000
  elif any(x in text for x in ["dua orang", "berboncengan"]):
    return 250000
  elif any(
      x in text
      for x in [
          "rambu",
          "marka",
          "lampu",
          "arus",
          "jalur",
          "ganjil",
          "genap",
          "kecepatan",
          "stnk",
          "keabsahan",
      ]
  ):
    return 500000
  else:
    return 0


# --- SCRAPE DETAIL PELANGGARAN VIA HTTP ---
def scrape_violation_details(session, detail_url):
  detailed_data = {}
  try:
    r = session.get(detail_url, headers=HTTP_HEADERS, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    container = soup.select_one("div.col-10.d-flex.flex-column")

    if container:
      # Ambil Link Foto jika ada
      img_element = container.find("img")
      if img_element and img_element.get("src"):
        img_src = img_element["src"]
        full_url = (
            "https://etle-pmj.id" + img_src
            if img_src.startswith("/")
            else img_src
        )
        detailed_data["Path Foto"] = full_url
      else:
        detailed_data["Path Foto"] = "-"

      full_text = container.get_text()

      def extract_info(text, keyword):
        try:
          return text.split(keyword)[1].split("\n")[1].strip()
        except (IndexError, ValueError):
          return "-"

      detailed_data["Detail Jenis Pelanggaran"] = extract_info(
          full_text, "Data Pelanggaran"
      )
      detailed_data["Dasar Hukum"] = extract_info(
          full_text, detailed_data.get("Detail Jenis Pelanggaran", "")
      )
      detailed_data["Waktu Kejadian"] = extract_info(
          full_text, "Hari, Tanggal & Waktu"
      )
      detailed_data["Merk"] = extract_info(full_text, "Merk")
      detailed_data["Tipe"] = extract_info(full_text, "Tipe")
      detailed_data["Warna"] = extract_info(full_text, "Warna")
      detailed_data["Nomor Rangka Detail"] = extract_info(
          full_text, "Nomor Rangka"
      )
      detailed_data["Model"] = extract_info(full_text, "Model")
      detailed_data["Tahun"] = extract_info(full_text, "Tahun")
      detailed_data["Nomor Mesin Detail"] = extract_info(
          full_text, "Nomor Mesin"
      )
      detailed_data["STNK Berlaku Sampai"] = extract_info(
          full_text, "STNK Berlaku Sampai"
      )
  except Exception as e:
    print(f"    -> Gagal scrape detail URL: {e}")

  return detailed_data


# --- CEK STATUS UTAMA VIA HTTP ---
def check_etle_status(session, url):
  try:
    r = session.get(url, headers=HTTP_HEADERS, timeout=20)
    soup = BeautifulSoup(r.text, "html.parser")

    # Cek Pop-up alert jika data tidak ditemukan
    popup = soup.select_one("div.swal2-html-container")
    if popup and "tidak ditemukan" in popup.text.lower():
      return "Aman", None

    # Parse baris tabel HTML
    rows = soup.select("tbody tr")
    violations = []
    for row in rows:
      cols = row.find_all("td")
      detail_link = ""
      link_element = row.select_one("a.btn-secondary")
      if link_element and link_element.get("href"):
        href = link_element["href"]
        if href.startswith("http"):
          detail_link = href
        else:
          detail_link = (
              "https://etle-pmj.id" + href
              if href.startswith("/")
              else "https://etle-pmj.id/" + href
          )

      if len(cols) >= 4:
        violations.append({
            "Tanggal": cols[0].get_text(strip=True),
            "Lokasi": cols[1].get_text(strip=True),
            "Jenis Pelanggaran": cols[2].get_text(strip=True),
            "Status": cols[3].get_text(strip=True),
            "Link Detail": detail_link,
        })

    if len(violations) > 0:
      return "Ada ETLE", violations
    else:
      return "Aman", None
  except Exception as e:
    return f"Error: {e}", None


# --- ORCHESTRATOR PROSES EXCEL ---
def process_excel(input_path, output_path):
  print("--- Memulai Proses Cek ETLE via HTTP Request ---")
  try:
    df = pd.read_excel(input_path, dtype=str)

    # Cari kolom dinamis
    col_plat = next((c for c in df.columns if "plat" in c.lower()), None)
    col_rangka = next((c for c in df.columns if "rangka" in c.lower()), None)
    col_mesin = next((c for c in df.columns if "mesin" in c.lower()), None)

    if not col_plat or not col_rangka or not col_mesin:
      print("ERROR: Kolom Plat/Rangka/Mesin tidak ditemukan di Excel!")
      return False

    final_report = []
    session = requests.Session()

    for idx, row in df.iterrows():
      plat = str(row[col_plat]).strip()
      rangka = str(row[col_rangka]).strip()
      mesin = str(row[col_mesin]).strip()
      real_row = idx + 2

      if not plat or plat.lower() == "nan":
        continue

      print(f"[{idx+1}/{len(df)}] Cek Full: {plat}")
      clean_plat = plat.replace(" ", "").replace("-", "")
      params = {
          "aksi": "cek",
          "nopol": clean_plat,
          "norangka": rangka,
          "nomesin": mesin,
      }
      url = "https://etle-pmj.id/?" + urlencode(params)

      status, violations = check_etle_status(session, url)

      base_row = dict.fromkeys(FULL_HEADERS, "-")
      base_row.update({
          "No. Baris": real_row,
          "Plat Nomor": plat,
          "No Rangka": rangka,
          "No Mesin": mesin,
          "Status ETLE": status,
          "Link Pengecekan": url,
      })

      if status == "Ada ETLE" and violations:
        print(f"  -> DETECTED: {len(violations)} Pelanggaran")
        for v in violations:
          row_data = base_row.copy()
          row_data.update(v)

          if v.get("Link Detail"):
            print("     -> Scraping detail via HTTP...")
            det = scrape_violation_details(session, v["Link Detail"])
            row_data.update(det)

          if "Jenis Pelanggaran" in row_data:
            row_data["Estimasi Denda"] = hitung_denda(
                row_data["Jenis Pelanggaran"]
            )

          final_report.append(row_data)
      else:
        print(f"  -> {status}")
        final_report.append(base_row)

    # Save output sesuai format asli 24 kolom
    df_hasil = pd.DataFrame(final_report)
    df_hasil = df_hasil.reindex(columns=FULL_HEADERS)
    df_hasil.to_excel(output_path, index=False)
    print(f"Excel berhasil disimpan: {output_path}")
    return True

  except Exception as e:
    print(f"Gagal proses excel: {e}")
    return False


# --- LOGIKA EMAIL ---
def run_email_bot():
  print("Menghubungkan ke Gmail IMAP...")
  try:
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, '(UNSEEN SUBJECT "CEK ETLE")')
    email_ids = messages[0].split()

    if not email_ids:
      print("Zzz... Tidak ada email baru dengan subjek 'CEK ETLE'.")
      return

    print(f"Ada {len(email_ids)} email baru!")

    for e_id in email_ids:
      res, msg_data = mail.fetch(e_id, "(RFC822)")
      for response_part in msg_data:
        if isinstance(response_part, tuple):
          msg = email.message_from_bytes(response_part[1])
          sender = msg.get("From")
          subject = decode_header(msg["Subject"])[0][0]
          if isinstance(subject, bytes):
            subject = subject.decode()

          print(f"\nMemproses email dari: {sender}")

          input_filename = "temp_input.xlsx"
          has_attachment = False
          for part in msg.walk():
            if part.get_content_maintype() == "multipart":
              continue
            if part.get("Content-Disposition") is None:
              continue
            fname = part.get_filename()
            if fname and (fname.endswith(".xlsx") or fname.endswith(".xls")):
              with open(input_filename, "wb") as f:
                f.write(part.get_payload(decode=True))
              has_attachment = True
              print(f"Lampiran ditemukan: {fname}")
              break

          if has_attachment:
            output_filename = "HASIL_LENGKAP.xlsx"
            if process_excel(input_filename, output_filename):
              send_reply(sender, output_filename)
              # Clean up temporary files
              if os.path.exists(input_filename):
                os.remove(input_filename)
              if os.path.exists(output_filename):
                os.remove(output_filename)
          else:
            print("Email tidak memiliki lampiran Excel.")

    mail.close()
    mail.logout()
  except Exception as e:
    print(f"Error Email Connection: {e}")


def send_reply(to_email, attachment_path):
  print(f"Mengirim balasan ke {to_email}...")
  msg = MIMEMultipart()
  msg["From"] = EMAIL_USER
  msg["To"] = to_email
  msg["Subject"] = "RE: Hasil Cek ETLE (Detail Lengkap)"
  msg.attach(
      MIMEText(
          "Berikut hasil pengecekan ETLE lengkap dengan detail pelanggaran.",
          "plain",
      )
  )

  with open(attachment_path, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
  encoders.encode_base64(part)
  part.add_header(
      "Content-Disposition",
      f"attachment; filename= {os.path.basename(attachment_path)}",
  )
  msg.attach(part)

  server = smtplib.SMTP("smtp.gmail.com", 587)
  server.starttls()
  server.login(EMAIL_USER, EMAIL_PASS)
  server.send_message(msg)
  server.quit()
  print("Balasan berhasil terkirim!")


if __name__ == "__main__":
  if not EMAIL_USER or not EMAIL_PASS:
    print("ERROR: Secret EMAIL_USER atau EMAIL_PASS belum diset!")
  else:
    run_email_bot()
