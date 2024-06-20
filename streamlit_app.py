from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
import io

replace_dict = {
                'Sekretariat Jenderal' : 1,
                'Inspektorat Jenderal' : 2,
                'Direktorat Jenderal Anggaran' : 3,
                'Direktorat Jenderal Pajak' : 4,
                'Direktorat Jenderal Bea dan Cukai' : 5,
                'Direktorat Jenderal Perbendaharaan' : 6,
                'Direktorat Jenderal Kekayaan Negara' : 7,
                'Direktorat Jenderal Perimbangan Keuangan' : 8,
                'Direktorat Jenderal Pengelolaan Pembiayaan Dan Risiko' : 9,
                'Direktorat Jenderal Pengelolaan Utang' : 10,
                'Badan Kebijakan Fiskal' : 11,
                'Badan Pengawas Pasar Modal dan Lembaga Keuangan' : 12,
                'Badan Pendidikan dan Pelatihan Keuangan' : 13,
                'Kementerian Agama' : 14,
                'Kementerian Agraria dan Tata Ruang/Badan Pertanahan Nasional' : 15,
                'Kementerian Badan Usaha Milik Negara' : 16,
                'Kementerian Dalam Negeri' : 17,
                'Kementerian Desa, Pembangunan Daerah Tertinggal, dan Transmigrasi' : 18,
                'Kementerian Energi dan Sumber Daya Mineral' : 19,
                'Kementerian Hukum dan Hak Asasi Manusia' : 20,
                'Kementerian Kelautan dan Perikanan' : 21,
                'Kementerian Kesehatan' : 22,
                'Kementerian Ketenagakerjaan' : 23,
                'Kementerian Komunikasi dan Informatika' : 24,
                'Kementerian Koordinator Bidang Kemaritiman dan Investasi' : 25,
                'Kementerian Koordinator Bidang Pembangunan Manusia dan Kebudayaan' : 26,
                'Kementerian Koordinator Bidang Perekonomian' : 27,
                'Kementerian Koordinator Bidang Politik, Hukum dan Keamanan' : 28,
                'Kementerian Koperasi dan Usaha Kecil dan Menengah' : 29,
                'Kementerian Lingkungan Hidup dan Kehutanan' : 30,
                'Kementerian Luar Negeri' : 31,
                'Kementerian Pariwisata dan Ekonomi Kreatif' : 32,
                'Kementerian Pekerjaan Umum dan Perumahan Rakyat' : 33,
                'Kementerian Pemberdayaan Perempuan dan Perlindungan Anak' : 34,
                'Kementerian Pemuda dan Olah Raga' : 35,
                'Kementerian Pendayagunaan Aparatur Negara dan Reformasi Birokrasi' : 36,
                'Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi' : 37,
                'Kementerian Perdagangan' : 38,
                'Kementerian Perencanaan Pembangunan Nasional / Badan Perencanaan Pembangunan Nasional' : 39,
                'Kementerian Perhubungan' : 40,
                'Kementerian Perindustrian' : 41,
                'Kementerian Pertahanan' : 42,
                'Kementerian Pertanian' : 43,
                'Kementerian Riset dan Teknologi / Badan Riset dan Inovasi Nasional' : 44,
                'Kementerian Sekretariat Negara' : 45,
                'Kementerian Sosial' : 46,
                'Kepolisian Negara Republik Indonesia' : 47,
                'Tentara Nasional Indonesia' : 48,
                'Tentara Nasional Indonesia Angkatan Darat' : 49,
                'Tentara Nasional Indonesia Angkatan Laut' : 50,
                'Tentara Nasional Indonesia Angkatan Udara' : 51,
                'Kejaksaan Republik Indonesia' : 52,
                'Mahkamah Agung' : 53,
                'Arsip Nasional Republik Indonesia' : 54,
                'Badan Ekonomi Kreatif' : 55,
                'Badan Informasi Geospasial' : 56,
                'Badan Intelijen Negara' : 57,
                'Badan Karantina Indonesia' : 58,
                'Badan Keamanan Laut' : 59,
                'Badan Kepegawaian Negara' : 60,
                'Badan Kependudukan dan Keluarga Berencana Nasional' : 61,
                'Badan Koordinasi Penanaman Modal' : 62,
                'Badan Meteorologi, Klimatologi dan Geofisika' : 63,
                'Badan Narkotika Nasional' : 64,
                'Badan Nasional Penanggulangan Bencana' : 65,
                'Badan Nasional Penanggulangan Terorisme' : 66,
                'Badan Nasional Pencarian dan Pertolongan' : 67,
                'Badan Nasional Penempatan dan Perlindungan Tenaga Kerja Indonesia' : 68,
                'Badan Nasional Pengelola Perbatasan' : 69,
                'Badan Pelindungan Pekerja Migran Indonesia' : 70,
                'Badan Pembinaan Ideologi Pancasila' : 71,
                'Badan Pemeriksa Keuangan' : 72,
                'Badan Penanggulangan Lumpur Sidoarjo' : 73,
                'Badan Pengawas Obat dan Makanan' : 74,
                'Badan Pengawas Pemilihan Umum' : 75,
                'Badan Pengawas Tenaga Nuklir' : 76,
                'Badan Pengawasan Keuangan dan Pembangunan' : 77,
                'Badan Pengembangan Wilayah Suramadu' : 78,
                'Badan Pengkajian dan Penerapan Teknologi' : 79,
                'Badan Pengusahaan Kawasan Perdagangan Bebas dan Pelabuhan Bebas Batam' : 80,
                'Badan Pengusahaan Kawasan Perdagangan Bebas dan Pelabuhan Bebas Sabang' : 81,
                'Badan Pusat Statistik' : 82,
                'Badan Riset dan Inovasi Nasional' : 83,
                'Badan Siber dan Sandi Negara' : 84,
                'Badan Standardisasi Nasional' : 85,
                'Badan Tenaga Nuklir Nasional' : 86,
                'Dewan Ketahanan Nasional' : 87,
                'Dewan Perwakilan Daerah' : 88,
                'Dewan Perwakilan Rakyat' : 89,
                'Komisi Nasional Hak Asasi Manusia' : 90,
                'Komisi Pemberantasan Korupsi' : 91,
                'Komisi Pemilihan Umum' : 92,
                'Komisi Pengawas Persaingan Usaha' : 93,
                'Komisi Yudisial Republik Indonesia' : 94,
                'Lembaga Administrasi Negara' : 95,
                'Lembaga Ilmu Pengetahuan Indonesia' : 96,
                'Lembaga Kebijakan Pengadaan Barang/Jasa Pemerintah' : 97,
                'Lembaga Ketahanan Nasional' : 98,
                'Lembaga National Single Window' : 99,
                'Lembaga Penerbangan dan Antariksa Nasional' : 100,
                'Lembaga Penyiaran Publik Radio Republik Indonesia' : 101,
                'Lembaga Penyiaran Publik Televisi Republik Indonesia' : 102,
                'Mahkamah Konstitusi RI' : 103,
                'Majelis Permusyawaratan Rakyat' : 104,
                'Ombudsman Republik Indonesia' : 105,
                'Otorita Ibu Kota Nusantara' : 106,
                'Perpustakaan Nasional Republik Indonesia' : 107,
                'Pusat Pelaporan dan Analisis Transaksi Keuangan' : 108,
                'Sekretariat Kabinet' : 109,
                'Pemprov/Pemkab/Pemda' : 110,
                'Non Kemenkeu' : 111

                }

"""
# Generate Pengumuman Kelulusan!

Download template dimari [template](https://docs.google.com/document/d/1CHTs9yKSYHYvJB4yrHhU1HlEjh_AaRHg/edit?usp=sharing&ouid=106976473759391912108&rtpof=true&sd=true)

Data kelulusan semantik - download melalui semantik menu kelulusan
Contoh file : student.graduate.template.20240614074537.xlsx

File bisa diupload sekaligus jika lebih dari satu kelas

"""
with st.form("my_form"):
    nama_pelatihan = st.text_input("Nama Pelatihan")
    tanggal_pelatihan = st.text_input("Tanggal Pelatihan","contoh: 22 s.d. 31 Agustus 2024")
    tempat_pelatihan = st.selectbox('Penyelenggara', ['Pusdiklat Anggaran dan Perbendaharaan',
                                                      'Balai Diklat Keuangan Medan',
                                                      'Balai Diklat Keuangan Pekanbaru',
                                                      'Balai Diklat Keuangan Palembang',
                                                      'Balai Diklat Keuangan Cimahi',
                                                      'Balai Diklat Keuangan Yogyakarta',
                                                      'Balai Diklat Keuangan Malang',
                                                      'Balai Diklat Keuangan Pontianak',
                                                      'Balai Diklat Keuangan Balikpapan',
                                                      'Balai Diklat Keuangan Denpasar',
                                                      'Balai Diklat Keuangan Makassar',
                                                      'Balai Diklat Keuangan Manado'])
    jenis_sertifikat = st.selectbox('Pilih status kelulusan', ['lulus','telah mengikuti'])
    template_pengumuman = st.file_uploader("Upload template pengumuman")
    data_kelulusan = st.file_uploader("Upload data kelulusan peserta semantik", accept_multiple_files=True)
    proses = st.form_submit_button('Proses')

if proses:
    document = MailMerge(template_pengumuman)
    document.merge(
        nama_pelatihan = nama_pelatihan,
        nama_pelatihan_upper = nama_pelatihan.upper(),
        tanggal_pelatihan = tanggal_pelatihan,
        tempat_pelatihan = tempat_pelatihan,
        tempat_pelatihan_upper = tempat_pelatihan.upper(),
        jenis_sertifikat = jenis_sertifikat,
        jenis_sertifikat_upper = jenis_sertifikat.upper()
    )
    df_all = pd.DataFrame()
    for excel in data_kelulusan:
        df = pd.read_excel(excel, sheet_name="Sheet1", skiprows=5, usecols='B:F')
        df_all = pd.concat([df_all,df])

    df_all.rename(columns = {'STATUS \nKELULUSAN':'STATUS'}, inplace = True)
    df_lulus=df_all[df_all["STATUS"] == 1]
    df_lulus['SATKER'] = df_lulus['SATKER'].str.split(' -', expand=True)[0]
    df_pengumuman = df_lulus[['NAMA','NIP','SATKER']].sort_values(by='NAMA', ascending=True)
    df_pengumuman['URUT'] = df_pengumuman['SATKER']
    df_pengumuman['URUT'].replace(replace_dict,inplace=True)
    df_dict = df_pengumuman[['NAMA','NIP','SATKER','URUT']].sort_values(by=['URUT', 'NAMA'], ascending=[True,True]).to_dict('records')
    document.merge_rows('NAMA', df_dict)
    document.write('template_word_.docx')
    doc = DocxTemplate('template_word_.docx')
    bio = io.BytesIO()
    doc.save(bio)
    if doc:
        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name="Pengumuman.docx",
            mime="docx"
        )
