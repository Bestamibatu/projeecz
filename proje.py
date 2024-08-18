import tkinter as tk
from tkinter import filedialog
import pandas as pd
from PIL import Image, ImageTk
import threading
import os
from datetime import datetime
import requests
import sys

class VerticalTabbarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vertical Tabbar with Logo")
        self.root.geometry("600x400")

        # Ana frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Sekme çubuğu için frame
        self.tabbar = tk.Frame(main_frame, width=200, bg='#2bb1e1')
        self.tabbar.pack(side=tk.LEFT, fill=tk.Y)

        # Sekme içerikleri için frame
        self.tab_contents = tk.Frame(main_frame, bg='#f7b900')
        self.tab_contents.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Logo
        self.load_logo()

        # Excel dosya yükleme ekranı
        self.excel_upload_screen = tk.Button(self.tabbar, text="Excel Yükleme Ekranı", command=self.show_excel_page, anchor='w', bg='#fff', fg='#d52d15', relief=tk.FLAT, font=("Arial", 12))
        self.excel_upload_screen.pack(fill=tk.X, pady=5)

        # Yarışmalar için collapse butonu ve ok işareti
        self.comp_button = tk.Button(self.tabbar, text="Yarışmalar ▼", command=self.toggle_comp_page, anchor='w', bg='#2bb1e1', relief=tk.FLAT, font=("Arial", 12))
        self.comp_button.pack(fill=tk.X)

        # Yarışmaların içeriği ve butonlar
        self.comp_frame = tk.Frame(self.tabbar)
        self.comp_frame.pack(fill=tk.X)

        self.comp_content_frame = tk.Frame(self.comp_frame, bg='#f7b900')
        self.comp_content_frame.pack(fill=tk.X)

        self.comp_buttons = []
        self.df = None

        self.update_comp_buttons()

        # Dosya adı etiketini tanımla
        self.selected_label = tk.Label(self.tab_contents, text="Seçilen Dosya: ", bg='#f7b900', font=("Arial", 14))
        self.selected_label.pack()

        self.clear_contents()



    def excel_select_file(self):
        # Dosya seçme penceresi açma
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Excel Dosyasını Seç"
        )
        if file_path:
            # Sadece dosya adını al ve etikete ekle
            file_name = os.path.basename(file_path)
            self.selected_label.config(text=f"Seçilen Dosya: {file_name}")  # Etiketi güncelle
            self.df = file_path
        else:
            self.selected_label.config(text="Seçilen Dosya: Hiçbir dosya seçilmedi.")  # Etiketi güncelle
        
    def load_logo(self):
        # Kaynak dosyanın yolunu almak için get_resource_path fonksiyonunu kullan
        if getattr(sys, 'frozen', False):
            # PyInstaller ile paketlendiğinde
            base_path = sys._MEIPASS
        else:
            # Geliştirme aşamasında
            base_path = os.path.dirname(os.path.abspath(__file__))

        logo_path = os.path.join(base_path, "t3vakfi.png")

        if os.path.isfile(logo_path):
            # Resmi yükle
            image = Image.open(logo_path).convert("RGBA")

            # Resmi yeniden boyutlandır
            max_width = 180  # Maksimum genişlik
            max_height = 80  # Maksimum yükseklik
            image.thumbnail((max_width, max_height))

            # Resmi ImageTk formatına dönüştür
            photo = ImageTk.PhotoImage(image)

            # Logo'yu bir label içinde göster
            self.logo = tk.Label(self.tabbar, image=photo, bg='#2bb1e1')
            self.logo.photo = photo  # Referansı sakla, aksi takdirde resim görünmeyebilir
            self.logo.pack(pady=10)
        else:
            print(f"Logo dosyası bulunamadı: {logo_path}")


            
    def load_excel(self):
        # Excel dosyasını yüklemek için dosya seçme penceresi
        file_path = self.df
        #file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            # Veriyi okuma işlemini ayrı bir iş parçacığında yapıyoruz
            threading.Thread(target=self.read_excel, args=(file_path,)).start()

    def read_excel(self, file_path):
        try:
            self.df = pd.read_excel(file_path)
            print(self.df)  # Verileri konsola yazdır

            # Klasörleri oluştur
            self.create_folders_and_files()

            # GUI'yi güncellemek için ana iş parçacığına dön
            self.root.after(0, self.update_comp_buttons)
            self.root.after(0, self.show_excel_page)
        except Exception as e:
            print(f"Error loading Excel file: {e}")

    def create_folders_and_files(self):
        if self.df is None or self.df.empty:
            return

        # Yıl bilgisini al
        now = datetime.now()
        year_str = now.strftime("%Y")

        # Proje dizinini al
        base_folder = self.get_resource_path("Dosyalar")

        # İlk sütunun adını al
        if not self.df.empty:
            first_column_name = self.df.columns[0]  # İlk sütun adı
            
            # Her yarışma için klasör oluştur
            for value in self.df[first_column_name].dropna().unique():
                competition_folder = os.path.join(base_folder, f"{year_str}/{value}")
                try:
                    os.makedirs(competition_folder, exist_ok=True)
                    print(f"Created folder: {competition_folder}")

                    # Her yarışma için ilgili takımların isimlerini ve URL'leri içeren dosyaları oluştur
                    if len(self.df.columns) > 3:
                        team_column_name = self.df.columns[1]
                        team_id_column_name = self.df.columns[2]
                        url_column_name = self.df.columns[3]
                        teams_urls = self.df[self.df[first_column_name] == value][[team_column_name, team_id_column_name, url_column_name]].dropna()
                        
                        for _, row in teams_urls.iterrows():
                            team = row[team_column_name]
                            team_id = row[team_id_column_name]
                            url = row[url_column_name]

                            # Takım ismi ve ID'si ile dosya yolu oluştur
                            file_name = f"{team}_{team_id}.pdf"
                            file_path = os.path.join(competition_folder, file_name)
                            
                            # PDF dosyasını indir
                            self.download_pdf(url, file_path)
                    
                except Exception as e:
                    print(f"Error creating folder {competition_folder}: {e}")

    def download_pdf(self, url, file_path):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                with open(file_path, 'wb') as file:
                    file.write(response.content)
                print(f"Downloaded PDF: {file_path}")
            else:
                print(f"Failed to download {url}: HTTP {response.status_code}")
        except Exception as e:
            print(f"Error downloading PDF from {url}: {e}")

    def update_comp_buttons(self):
        # Yarışma butonlarını oluşturmak için dosya sistemini kontrol et
        base_folder = self.get_resource_path("Dosyalar")
    
        # Önceki butonları kaldır
        for widget in self.comp_content_frame.winfo_children():
            widget.destroy()

        # Klasörlerdeki yarışma isimlerini al
        if os.path.exists(base_folder):
            year_str = datetime.now().strftime("%Y")
            year_folder = os.path.join(base_folder, year_str)
            if os.path.exists(year_folder):
                for competition in os.listdir(year_folder):
                    competition_folder = os.path.join(year_folder, competition)
                    if os.path.isdir(competition_folder):
                        # Takımların sayısını hesapla
                        team_files = [f for f in os.listdir(competition_folder) if os.path.isfile(os.path.join(competition_folder, f))]
                        num_teams = len(team_files)
                    
                        # Butonu oluştur
                        button = tk.Button(
                            self.comp_content_frame,
                            text=f"Yarışma: {competition} ({num_teams} takım)",
                            command=lambda competition=competition: self.show_comp_page(competition),
                            bg='#2bb1e1',
                            fg='#fff',
                            relief=tk.FLAT,
                            font=("Arial", 12)
                        )
                        button.pack(fill=tk.X)
                        self.comp_buttons.append(button)


    def toggle_comp_page(self):
        # Yarışmalar sayfasını açma veya kapama
        if self.comp_frame.winfo_ismapped():
            self.comp_frame.pack_forget()
            self.comp_button.config(text="Yarışmalar ▼")
        else:
            self.comp_frame.pack(fill=tk.X)
            self.comp_button.config(text="Yarışmalar ▲")

    def show_comp_page(self, competition):
        # Yarışma içeriğini göstermek için
        self.clear_contents()
        comp_content = tk.Label(self.tab_contents, text=f"Yarışma: {competition}", bg='#f7b900', font=("Arial", 16))
        comp_content.pack(padx=10, pady=10)

        # Excel verilerinin ikinci sütununu göster
        base_folder = self.get_resource_path("Dosyalar")
        year_str = datetime.now().strftime("%Y")
        competition_folder = os.path.join(base_folder, year_str, competition)

        if os.path.exists(competition_folder):
            team_files = [f for f in os.listdir(competition_folder) if os.path.isfile(os.path.join(competition_folder, f))]
            if len(team_files) > 0:
                table_frame = tk.Frame(self.tab_contents, bg='#f7b900')
                table_frame.pack(pady=10)
                tk.Label(table_frame, text="Takımlar ve Dosyalar", bg='#f7b900', font=("Arial", 18)).pack()
            for team_file in team_files:
                team_name = os.path.splitext(team_file)[0]  # Dosya adından uzantıyı çıkar
                # Buton oluştur ve tıklama olayına dosyayı açma fonksiyonunu ata
                button = tk.Button(table_frame, text=team_name, bg='#f7b900', font=("Arial", 12), command=lambda file_path=os.path.join(competition_folder, team_file): self.open_file(file_path))
                button.pack(pady=2)

    def open_file(self, file_path):
        # Dosyayı varsayılan uygulamada aç
        try:
            os.startfile(file_path)
        except Exception as e:
            print(f"Error opening file {file_path}: {e}")
            
    def apply_button_styles(self):
        # Butonun kenarlarına gölge efekti
        self.stylish_button.bind("<Enter>", self.on_enter)
        self.stylish_button.bind("<Leave>", self.on_leave)

    def show_excel_page(self):
        # Excel verilerini gösteren sayfayı aç
        self.clear_contents()
        upload_label = tk.Label(self.tab_contents, text="Dosya Yükleme", bg='#f7b900', font=("Arial", 16))
        upload_label.pack(padx=10, pady=10)
        select_button = tk.Button(self.tab_contents,
            command=self.excel_select_file,
            text="Excel Dosyasını Seç",
            font=("Arial", 14, "bold"),
            bg='#2bb1e1',
            fg='#fff',
            relief=tk.FLAT,
            bd=0,
            padx=15,  # İç boşluk (sol ve sağ)
            pady=5,   # İç boşluk (üst ve alt)
            highlightthickness=0,
            width=15   # Butonun genişliği (karakter sayısı cinsinden)
        )
        select_button.pack(pady=20, anchor='center')
        self.selected_label = tk.Label(self.tab_contents, text="Seçilen Dosya: ", bg='#f7b900', font=("Arial", 14))
        self.selected_label.pack()
        upload_button = tk.Button(self.tab_contents,
            text="Dosyayı Yükle",
            command=self.load_excel,
            font=("Arial", 14, "bold"),
            bg='#2bb1e1',
            fg='#fff',
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=10,
            highlightthickness=0
        )
        upload_button.pack(pady=5)

        excel_page = tk.Label(self.tab_contents, text="Excel Dosyasından Alınan Yarışmalar", bg='#f7b900', font=("Arial", 16))
        excel_page.pack(padx=10, pady=10)

        # Eğer Excel verileri yüklendiyse, ilk sütundaki verileri ekleyin
        if self.df is not None and len(self.df.columns) > 0:
            first_column_name = self.df.columns[0]
            added_values = set()  # Eklenecek olan yarışma isimlerini tutacak set
            
            for value in self.df[first_column_name].dropna():
                if value not in added_values:
                    data_label = tk.Label(self.tab_contents, text=value, bg='#f7b900', font=("Arial", 12))
                    data_label.pack(pady=2)
                    added_values.add(value)  # Ekleme yapıldıktan sonra set'e ekle

    def clear_contents(self):
        # İçerikleri temizleme
        for widget in self.tab_contents.winfo_children():
            widget.pack_forget()

    def get_resource_path(self, relative_path):
        """ Return the path to the resource, handling PyInstaller's temp path if needed """
        if getattr(sys, 'frozen', False):
        # PyInstaller ile paketlendiğinde
            base_path = os.path.dirname(sys.executable)
        else:
        # Geliştirme aşamasında
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = VerticalTabbarApp(root)
    root.mainloop()
