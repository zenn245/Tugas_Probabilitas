import os
import sys
import warnings

# Mengabaikan pesan peringatan minor dari format Excel bawaan
warnings.filterwarnings("ignore")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

# ==============================================================================
# 1. MEMBACA FILE EXCEL SECARA DINAMIS
# ==============================================================================
# Mendapatkan direktori folder tempat file skrip ini berada
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
file_excel = os.path.join(BASE_DIR, "LATIHAN -PROBABILITAS_Zain.xlsx")

if not os.path.exists(file_excel):
  print(f"[ERROR] File tidak ditemukan di: {file_excel}")
  print(
      "Pastikan file Excel berada di satu folder yang sama dengan file skrip"
      " ini."
  )
  sys.exit()

# Pada data Excel Anda, header tabel dimulai dari baris ke-7 (skiprows=6)
df_raw = pd.read_excel(file_excel, sheet_name=0, skiprows=6)

# Mengambil kolom data mahasiswa (ID, Name, Quiz, Mid, Final, Final Score, Grade)
df = df_raw.iloc[:, 1:8].copy()
df.columns = ["ID", "Nama", "Quiz", "Mid_Term", "Final_Term", "Final_Score", "Grade"]

# Membersihkan baris kosong dan baris catatan di bawah tabel
df = df.dropna(subset=["ID", "Nama"]).reset_index(drop=True)

# Memastikan tipe data nilai berupa angka (numeric)
kolom_nilai = ["Quiz", "Mid_Term", "Final_Term", "Final_Score"]
for col in kolom_nilai:
  df[col] = pd.to_numeric(df[col], errors="coerce")

print("=" * 65)
print(f" SUKSES: Data berhasil dimuat ({len(df)} Mahasiswa)")
print("=" * 65)
print(df.head(6))


# ==============================================================================
# 2. STATISTIKA DESKRIPTIF
# ==============================================================================
print("\n" + "=" * 65)
print(" STATISTIKA DESKRIPTIF KOMPONEN NILAI")
print("=" * 65)

stats = df[kolom_nilai].describe().T
stats["median"] = df[kolom_nilai].median()
stats["variance"] = df[kolom_nilai].var()

# Tampilkan ringkasan metrik deskriptif
tabel_stats = stats[
    ["mean", "median", "std", "variance", "min", "25%", "50%", "75%", "max"]
].round(2)
print(tabel_stats)

print("\nDistribusi Frekuensi Grade:")
grade_counts = df["Grade"].value_counts()
print(grade_counts)


# ==============================================================================
# 3. VISUALISASI DATA (4 GRAFIK DALAM 1 TAMPILAN)
# ==============================================================================
sns.set_theme(style="whitegrid")
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle(
    "Visualisasi Data & Statistika Deskriptif Nilai Mahasiswa",
    fontsize=16,
    fontweight="bold",
    y=0.98,
)

# Plot 1: Bar Chart Distribusi Grade
urutan_grade = ["A", "AB", "B", "BC", "C", "D", "F"]
urutan_grade = [g for g in urutan_grade if g in df["Grade"].values]

sns.countplot(
    ax=axes[0, 0],
    data=df,
    x="Grade",
    order=urutan_grade,
    color="#3b82f6",
    edgecolor="black",
)
axes[0, 0].set_title(
    "1. Distribusi Frekuensi Grade Mahasiswa", fontweight="bold"
)
axes[0, 0].set_xlabel("Grade", fontweight="bold")
axes[0, 0].set_ylabel("Jumlah Mahasiswa", fontweight="bold")

# Menambahkan angka di atas bar diagram
for p in axes[0, 0].patches:
  val = int(p.get_height())
  axes[0, 0].annotate(
      f"{val}",
      (p.get_x() + p.get_width() / 2.0, val),
      ha="center",
      va="bottom",
      fontweight="bold",
      xytext=(0, 2),
      textcoords="offset points",
  )

# Plot 2: Histogram Nilai Akhir (Final Score) + Kurva Distribusi
sns.histplot(
    ax=axes[0, 1],
    data=df["Final_Score"],
    kde=True,
    color="#1d4ed8",
    bins=8,
    edgecolor="black",
)
mean_val = df["Final_Score"].mean()
median_val = df["Final_Score"].median()
axes[0, 1].axvline(
    mean_val,
    color="red",
    linestyle="--",
    linewidth=2,
    label=f"Mean: {mean_val:.2f}",
)
axes[0, 1].axvline(
    median_val,
    color="green",
    linestyle="-.",
    linewidth=2,
    label=f"Median: {median_val:.2f}",
)
axes[0, 1].set_title(
    "2. Distribusi Nilai Akhir (Final Score)", fontweight="bold"
)
axes[0, 1].set_xlabel("Nilai Akhir", fontweight="bold")
axes[0, 1].set_ylabel("Frekuensi", fontweight="bold")
axes[0, 1].legend()

# Plot 3: Box Plot Komponen Nilai (Melihat Sebaran & Outlier)
sns.boxplot(
    ax=axes[1, 0],
    data=df[kolom_nilai],
    palette="Set2",
    width=0.5,
    showmeans=True,
    meanprops={
        "marker": "o",
        "markerfacecolor": "red",
        "markeredgecolor": "black",
    },
)
axes[1, 0].set_title(
    "3. Sebaran Komponen Nilai (Titik Merah: Rata-rata)", fontweight="bold"
)
axes[1, 0].set_xticklabels(
    ["Quiz", "Mid Term", "Final Term", "Final Score"], fontweight="bold"
)
axes[1, 0].set_ylabel("Skor", fontweight="bold")

# Plot 4: Heatmap Korelasi Antar Komponen Nilai
matriks_korelasi = df[kolom_nilai].corr()
sns.heatmap(
    ax=axes[1, 1],
    data=matriks_korelasi,
    annot=True,
    fmt=".2f",
    cmap="Blues",
    cbar=True,
    square=True,
    linewidths=1,
)
axes[1, 1].set_title(
    "4. Matriks Korelasi Antara Komponen Nilai", fontweight="bold"
)

plt.tight_layout(rect=[0, 0, 1, 0.96])

# Menyimpan hasil visualisasi grafik ke format file PNG
gambar_output = os.path.join(BASE_DIR, "visualisasi_hasil.png")
plt.savefig(gambar_output, dpi=300)
print(f"\n[OK] Gambar grafik berhasil disimpan: '{gambar_output}'")

# Menampilkan grafik di layar
plt.show()


# ==============================================================================
# 4. MENYIMPAN TABEL EXCEL BERSIH (UNTUK TUGAS UPLOAD)
# ==============================================================================
excel_output = os.path.join(BASE_DIR, "tabel_hasil_analisis.xlsx")
with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
  df.to_excel(writer, sheet_name="Data_Mahasiswa", index=False)
  tabel_stats.to_excel(writer, sheet_name="Statistika_Deskriptif")

print(f"[OK] File Excel hasil analisis berhasil dibuat: '{excel_output}'")
print("\nSemua proses selesai dengan sukses tanpa error!")