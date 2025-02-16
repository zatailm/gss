import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import random
import pandas as pd
import time
import re
from concurrent.futures import ThreadPoolExecutor
from scholarly import scholarly

# Variabel global untuk status scraping
stop_scraping = threading.Event()
all_publications = []

def is_valid_author_name(name):
    """Validasi nama penulis hanya boleh huruf, angka, titik, koma, dan spasi."""
    return bool(re.match(r'^[a-zA-Z0-9., ]+$', name))

def show_info():
    """Menampilkan informasi tentang delay scraping."""
    messagebox.showinfo("Informasi", 
                        "Aplikasi ini menggunakan delay acak antara 0.5 hingga 2.5 detik, "
                        "sehingga proses scraping akan terasa lambat. Hal ini dilakukan untuk "
                        "mengurangi risiko pemblokiran oleh Google Scholar.\n\n"
                        "Aplikasi ini tidak menerapkan pergantian User-Agent secara acak/berkala "
                        "sehingga risiko pemblokiran masih tetap ada.\n\n"
                        "Jika mengalami pemblokiran, gunakan VPN atau IP yang berbeda.\n\n"
                        "Jika Anda ingin mensitasi perangkat lunak ini, gunakan: \n"
                        "Ilmam, M.A.Z. (2025). Google Scholar Scraper (Versi 1.0) [Perangkat Lunak Komputer].\n\n"
                        "Google Scholar Scraper v1.0 \n"
                        "© 2025 M. Adib Zata Ilmam")

def scrape_author_publications(author_name):
    """Scrape publikasi dari Google Scholar berdasarkan nama penulis."""
    global all_publications

    search_query = scholarly.search_author(author_name)
    author = next(search_query, None)

    if author is None:
        log_text.insert(tk.END, f"⚠ Tidak ditemukan: {author_name}\n\n")
        root.update_idletasks()
        return

    author = scholarly.fill(author)
    log_text.insert(tk.END, f"📌 Ditemukan: {author_name}\nMemulai scraping...\n")
    root.update_idletasks()

    for pub in author.get('publications', []):
        if stop_scraping.is_set():
            log_text.insert(tk.END, f"⛔ Scraping {author_name} dihentikan!\n\n")
            root.update_idletasks()
            return
        
        try:
            pub_details = scholarly.fill(pub)
            title = pub_details.get('bib', {}).get('title', 'N/A')
            authors = pub_details.get('bib', {}).get('author', 'N/A')
            citations = pub_details.get('num_citations', 0)
            year = pub_details.get('bib', {}).get('pub_year', 'N/A')

            # Ekstrak link publikasi
            link_pattern = r'https?://\S+'
            link_matches = re.findall(link_pattern, str(pub_details))
            link = link_matches[0] if link_matches else 'N/A'

            all_publications.append({
                'Pencarian': author_name, 
                'Judul': title, 
                'Nama Penulis': authors, 
                'Jumlah Sitasi': citations, 
                'Tahun Terbit': year, 
                'Link': link
            })

            log_text.insert(tk.END, f"✔ {author_name}: {title[:40]}... ({year})\n")
            log_text.see(tk.END)  
            root.update_idletasks()

            # Tambahkan delay acak 
            time.sleep(random.uniform(0.5, 2.5))

        except Exception as e:
            print(f"Error: {e}")
            continue

    log_text.insert(tk.END, f"✅ Scraping selesai untuk {author_name}!\n\n")
    root.update_idletasks()

def save_to_excel():
    """Simpan hasil scraping ke file Excel."""
    if not all_publications:
        messagebox.showerror("Error", "Tidak ada data untuk disimpan.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    if not file_path:
        return

    df = pd.DataFrame(all_publications)
    
    try:
        df.to_excel(file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Sukses", f"Data disimpan di:\n{file_path}")
        log_text.insert(tk.END, f"📁 File disimpan di: {file_path}\n")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan file!\n{e}")

def start_scraping_thread():
    """Mulai scraping di thread terpisah agar UI tetap responsif."""
    global all_publications
    stop_scraping.clear()
    all_publications = []

    author_names = author_entry.get().strip()
    if not author_names:
        messagebox.showerror("Error", "Masukkan minimal satu pencarian!")
        return

    author_list = [name.strip() for name in author_names.split(";") if name.strip()]

    # Validasi setiap nama penulis
    for author in author_list:
        if not is_valid_author_name(author):
            messagebox.showerror("Error", f"Nama penulis tidak valid: {author}")
            return

    log_text.delete(1.0, tk.END)  # Bersihkan log sebelum scraping baru
    progress_bar.start(10)
    scrape_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)

    def run_scraper():
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(scrape_author_publications, author) for author in author_list]

        # Aktifkan kembali tombol setelah selesai
        progress_bar.stop()
        scrape_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)
        log_text.insert(tk.END, "🔹 Scraping selesai! Tekan 'Simpan ke Excel' untuk menyimpan hasil.\n")

    threading.Thread(target=run_scraper, daemon=True).start()

def stop_scraping_process():
    """Hentikan scraping."""
    stop_scraping.set()
    stop_button.config(state=tk.DISABLED)

# ======================= GUI (Tkinter) =======================

root = tk.Tk()
root.title("Google Scholar Scraper 1.0")
root.geometry("500x550")
root.minsize(500, 550)  
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")

# Frame Utama
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill="both", expand=True)

# Tombol Informasi Kecil di Pojok Kanan Atas
info_button = tk.Button(root, text="🛈", font=("Arial", 9, "bold"), 
                        command=show_info, relief="flat")
info_button.place(x=462, y=10)  # Posisi di pojok kanan atas

# Label Pencarian
ttk.Label(main_frame, text="Masukkan Nama Penulis (pisahkan dengan ; jika lebih dari 1:)", font=("Arial", 10)).pack(pady=5)

# Entry Pencarian
author_entry = ttk.Entry(main_frame, width=50)
author_entry.pack(pady=5)

# Progress Bar
progress_bar = ttk.Progressbar(main_frame, length=400, mode="indeterminate")
progress_bar.pack(pady=5)

# Log Output
log_frame = ttk.Frame(main_frame)
log_frame.pack(fill="both", expand=True, pady=5)

log_text = tk.Text(log_frame, wrap="word", state=tk.NORMAL, height=15, font=("Consolas", 10))
log_text.pack(side="left", fill="both", expand=True)

scrollbar = ttk.Scrollbar(log_frame, command=log_text.yview)
scrollbar.pack(side="right", fill="y")

log_text.config(yscrollcommand=scrollbar.set)
log_text.insert(tk.END, "🔹 Log aktivitas akan tampil di sini...\n")

# Tombol untuk Scraping & Simpan
button_frame = ttk.Frame(main_frame)
button_frame.pack(fill="x", pady=(5, 10))

scrape_button = ttk.Button(button_frame, text="Mulai Scraping", command=start_scraping_thread)
scrape_button.pack(side="left", padx=5, fill="x", expand=True)

stop_button = ttk.Button(button_frame, text="Stop", command=stop_scraping_process, state=tk.DISABLED)
stop_button.pack(side="left", padx=5, fill="x", expand=True)

save_button = ttk.Button(button_frame, text="Simpan ke Excel", command=save_to_excel)
save_button.pack(side="right", padx=5, fill="x", expand=True)

root.mainloop()

# build
# pyinstaller --onefile --noconsole --hidden-import=openpyxl.cell.cell --hidden-import=openpyxl.workbook.workbook gss.py
