import sqlite3
import openpyxl

DB_FILE = "musteri_kayit.db"

def db_baglanti():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""CREATE TABLE IF NOT EXISTS musteriler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ad_soyad TEXT NOT NULL,
        telefon TEXT,
        email TEXT,
        adres TEXT
    )""")
    return conn

def musteri_ekle(conn):
    ad = input("Ad Soyad: ")
    tel = input("Telefon: ")
    email = input("E-posta: ")
    adres = input("Adres: ")
    conn.execute("INSERT INTO musteriler (ad_soyad, telefon, email, adres) VALUES (?, ?, ?, ?)",
                 (ad, tel, email, adres))
    conn.commit()
    print("Müşteri eklendi.")

def musterileri_listele(conn):
    for row in conn.execute("SELECT * FROM musteriler"):
        print(row)

def musteri_ara(conn):
    kelime = input("Arama: ")
    for row in conn.execute("""SELECT * FROM musteriler 
                                WHERE ad_soyad LIKE ? OR telefon LIKE ? OR email LIKE ? OR adres LIKE ?""",
                            (f"%{kelime}%",)*4):
        print(row)

def musteri_sil(conn):
    id_ = input("Silinecek ID: ")
    conn.execute("DELETE FROM musteriler WHERE id=?", (id_,))
    conn.commit()
    print("Müşteri silindi.")

def musteri_guncelle(conn):
    id_ = input("Güncellenecek ID: ")
    ad = input("Yeni Ad Soyad: ")
    tel = input("Yeni Telefon: ")
    email = input("Yeni E-posta: ")
    adres = input("Yeni Adres: ")
    conn.execute("""UPDATE musteriler SET ad_soyad=?, telefon=?, email=?, adres=? WHERE id=?""",
                 (ad, tel, email, adres, id_))
    conn.commit()
    print("Müşteri güncellendi.")

def excel_aktar(conn):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Müşteriler"
    ws.append(["ID", "Ad Soyad", "Telefon", "E-posta", "Adres"])
    for row in conn.execute("SELECT * FROM musteriler"):
        ws.append(row)
    dosya_adi = "musteriler.xlsx"
    wb.save(dosya_adi)
    print(f"Excel'e aktarıldı: {dosya_adi}")

def menu():
    conn = db_baglanti()
    while True:
        print("\n1) Müşteri Ekle\n2) Listele\n3) Ara\n4) Sil\n5) Güncelle\n6) Excel'e Aktar\n7) Çıkış")
        secim = input("Seçim: ")
        if secim == "1": musteri_ekle(conn)
        elif secim == "2": musterileri_listele(conn)
        elif secim == "3": musteri_ara(conn)
        elif secim == "4": musteri_sil(conn)
        elif secim == "5": musteri_guncelle(conn)
        elif secim == "6": excel_aktar(conn)
        elif secim == "7": break
        else: print("Hatalı seçim.")

if __name__ == "__main__":
    menu()
