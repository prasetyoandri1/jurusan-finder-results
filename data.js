function seedDefaults_(){
  const ss=ss_(), ms=ss.getSheetByName('Majors'), as=ss.getSheetByName('Archetypes'), st=ss.getSheetByName('Settings');
  if(st.getLastRow()<2)st.getRange(2,1,5,3).setValues([
    ['form_id',JF.FORM_ID,'Tally form ID'],['brand_handle','@andr.selfedu','CTA sosial'],['brand_link','https://lynk.id/andr','CTA resource'],['score_scale','auto','Mendeteksi 0–10 atau 0–100'],['retention_days','730','Hapus raw payload lebih cepat bila diperlukan']]);
  if(as.getLastRow()<2)as.getRange(2,1,ARCHETYPE_ROWS.length,18).setValues(ARCHETYPE_ROWS);
  if(ms.getLastRow()<2)ms.getRange(2,1,MAJOR_ROWS.length,24).setValues(MAJOR_ROWS);
}

const ARCHETYPE_ROWS = [
 ['analyst','The Analyst','Menemukan pola, menguji bukti, menarik kesimpulan','Kuat dalam logika, data, dan pemecahan masalah',95,85,85,75,55,45,30,45,35,40,55,45,35,45],
 ['strategist','The Strategist','Melihat sistem, peluang, dan langkah berikutnya','Menggabungkan analisis, bisnis, dan kepemimpinan',80,70,65,45,55,60,45,70,55,85,95,65,35,55],
 ['innovator','The Innovator','Membangun solusi baru lewat teknologi dan eksperimen','Kuat pada teknologi, rekayasa, dan kreativitas terapan',70,70,75,75,95,90,70,45,35,55,55,30,35,60],
 ['creator','The Creator','Mengubah ide menjadi cerita, visual, dan pengalaman','Kuat pada kreativitas, komunikasi, dan ekspresi',40,25,45,30,40,35,100,90,65,45,60,30,30,45],
 ['connector','The Connector','Memahami orang dan membangun hubungan bermakna','Kuat pada komunikasi, sosial, dan kolaborasi',40,25,45,30,30,25,65,95,95,70,65,45,55,45],
 ['leader','The Leader','Menggerakkan orang menuju hasil','Kuat pada pengaruh, keputusan, dan eksekusi',55,45,40,30,35,35,50,85,80,100,90,70,40,45],
 ['advocate','The Advocate','Menjaga keadilan dan memperbaiki masyarakat','Kuat pada hukum, kebijakan, komunikasi, dan dampak sosial',60,35,70,35,30,25,50,85,90,75,60,100,45,60],
 ['healer','The Healer','Meningkatkan kesehatan manusia dan lingkungan','Kuat pada sains, kepedulian, dan pelayanan',55,45,75,90,45,35,30,60,85,55,35,35,100,80]
];

// id, major, cluster, active, min, wildcard, description, consideration, 14 weights, exclusion
const MAJOR_ROWS = [
 ['data-science','Sains Data','Data & Komputasi',true,62,false,'Mengolah data menjadi prediksi dan keputusan.','Intensif matematika, statistik, dan coding.',95,100,90,75,90,55,35,45,30,45,65,25,25,45,'NUM<45:25;ANL<45:20'],
 ['statistics','Statistika','Data & Komputasi',true,62,false,'Mendesain analisis data dan ketidakpastian.','Membutuhkan ketelitian dan matematika yang kuat.',95,100,95,80,65,35,25,40,30,35,55,25,35,45,'NUM<50:30'],
 ['informatics','Informatika','Data & Komputasi',true,60,false,'Membangun perangkat lunak dan sistem komputasi.','Butuh latihan coding konsisten.',85,85,75,65,100,70,55,40,25,40,55,20,20,35,'TEC<45:25'],
 ['information-systems','Sistem Informasi','Data & Komputasi',true,58,false,'Menghubungkan teknologi dengan proses bisnis.','Perlu nyaman dengan sistem dan komunikasi lintas tim.',75,65,60,45,90,50,45,70,45,65,85,25,20,35,'TEC<35:20'],
 ['cybersecurity','Keamanan Siber','Data & Komputasi',true,60,true,'Melindungi sistem, jaringan, dan data.','Menuntut ketelitian dan belajar teknologi terus-menerus.',90,80,85,60,100,65,35,40,30,45,45,45,20,35,'TEC<45:30'],
 ['ai','Kecerdasan Artifisial','Data & Komputasi',true,64,true,'Merancang sistem yang belajar dari data.','Sangat kuat di matematika, algoritma, dan coding.',95,100,95,85,100,70,50,40,25,45,55,20,25,45,'NUM<50:30;TEC<45:25'],
 ['computer-engineering','Teknik Komputer','Teknik',true,60,false,'Menggabungkan hardware, software, dan sistem tertanam.','Banyak matematika, fisika, dan praktikum.',80,85,70,80,100,95,40,35,25,45,45,20,25,40,'NUM<40:20;ENG<45:25'],
 ['electrical-engineering','Teknik Elektro','Teknik',true,60,false,'Mempelajari listrik, elektronika, kontrol, dan telekomunikasi.','Intensif matematika dan fisika.',80,90,70,90,85,100,35,35,25,45,40,20,25,50,'NUM<45:25;SCI<45:25'],
 ['mechanical-engineering','Teknik Mesin','Teknik',true,58,false,'Merancang mesin, energi, dan sistem manufaktur.','Banyak fisika, kalkulus, dan desain teknis.',75,80,65,90,75,100,45,35,25,45,45,20,25,55,'SCI<45:25'],
 ['civil-engineering','Teknik Sipil','Teknik',true,58,false,'Merancang infrastruktur dan lingkungan terbangun.','Perlu matematika, fisika, dan tanggung jawab keselamatan.',75,80,65,85,65,100,40,50,45,60,55,35,30,70,'NUM<40:20;ENG<40:20'],
 ['industrial-engineering','Teknik Industri','Teknik',true,58,false,'Mengoptimalkan sistem manusia, proses, dan teknologi.','Gabungan matematika, bisnis, dan operasi.',85,80,70,60,70,85,45,65,55,75,85,30,30,50,'ANL<40:20'],
 ['chemical-engineering','Teknik Kimia','Teknik',true,60,false,'Mengubah bahan melalui proses industri.','Kuat di kimia, matematika, dan keselamatan proses.',80,85,80,100,60,95,35,35,25,45,45,25,45,65,'SCI<50:30'],
 ['environmental-engineering','Teknik Lingkungan','Lingkungan',true,58,false,'Merancang solusi air, limbah, dan polusi.','Memadukan sains, rekayasa, dan kerja lapangan.',75,70,80,90,60,85,40,55,60,55,45,35,55,100,'ENV<40:20;SCI<35:20'],
 ['architecture','Arsitektur','Desain & Kreatif',true,58,false,'Merancang ruang yang fungsional dan bermakna.','Studio desain memakan waktu; perlu kreativitas dan ketekunan.',60,50,55,55,55,75,100,65,55,55,55,25,25,70,'CRT<45:25'],
 ['urban-planning','Perencanaan Wilayah dan Kota','Lingkungan',true,56,true,'Menyusun kota dan wilayah yang layak huni.','Memerlukan analisis spasial, kebijakan, dan komunikasi publik.',75,60,75,55,50,45,65,80,85,70,65,70,35,95,'ENV<35:20'],
 ['geodesy','Teknik Geodesi/Geomatika','Teknik',true,58,true,'Memetakan bumi dengan survei dan data spasial.','Butuh matematika, teknologi, dan aktivitas lapangan.',80,85,75,75,85,85,40,45,35,45,45,25,25,85,'NUM<40:20'],
 ['medicine','Kedokteran','Kesehatan',true,68,false,'Mendiagnosis dan menangani kesehatan manusia.','Pendidikan panjang, beban akademik tinggi, dan tanggung jawab besar.',80,65,95,100,45,45,30,75,90,65,35,35,100,45,'HLT<55:35;SCI<55:30'],
 ['dentistry','Kedokteran Gigi','Kesehatan',true,65,false,'Menangani kesehatan gigi dan mulut.','Butuh ketelitian tangan, sains, dan interaksi pasien.',70,55,80,95,45,65,45,75,85,55,40,30,100,40,'HLT<50:30;SCI<50:25'],
 ['pharmacy','Farmasi','Kesehatan',true,62,false,'Mempelajari obat, formulasi, dan penggunaan yang aman.','Kuat di kimia, biologi, dan ketelitian.',80,70,90,100,55,50,30,55,60,45,45,35,95,45,'SCI<50:30'],
 ['nursing','Keperawatan','Kesehatan',true,58,false,'Memberikan asuhan langsung dan berkelanjutan.','Perlu empati, stamina, dan kesiapan kerja klinis.',55,40,65,80,35,35,30,80,100,60,30,25,100,40,'HLT<50:30;SOC<45:25'],
 ['public-health','Kesehatan Masyarakat','Kesehatan',true,58,false,'Mencegah penyakit pada tingkat populasi.','Menggabungkan data, edukasi, kebijakan, dan kerja komunitas.',75,65,85,75,40,30,45,80,95,70,55,65,95,70,'HLT<40:20'],
 ['nutrition','Ilmu Gizi','Kesehatan',true,58,false,'Mempelajari pangan dan dampaknya pada kesehatan.','Memerlukan sains dan komunikasi perubahan perilaku.',65,55,75,90,35,35,45,75,85,50,45,30,95,55,'SCI<40:20;HLT<40:20'],
 ['psychology','Psikologi','Sosial & Humaniora',true,58,false,'Memahami perilaku dan proses mental manusia.','Banyak membaca, riset, statistik, dan batas etika profesi.',75,55,90,55,30,25,55,90,100,60,45,45,85,45,'SOC<45:25'],
 ['economics','Ilmu Ekonomi','Bisnis & Ekonomi',true,60,false,'Menganalisis pilihan, pasar, dan kebijakan.','Cukup kuantitatif dan model-driven.',95,90,85,55,45,30,35,65,65,65,90,65,30,60,'NUM<40:20'],
 ['management','Manajemen','Bisnis & Ekonomi',true,56,false,'Mengelola organisasi, strategi, pemasaran, dan operasi.','Spektrum luas; perlu pengalaman proyek agar menonjol.',65,55,50,30,40,30,55,85,75,90,100,45,30,40,'BUS<35:20'],
 ['accounting','Akuntansi','Bisnis & Ekonomi',true,58,false,'Menyusun dan memeriksa informasi keuangan.','Menuntut ketelitian, aturan, dan rutinitas angka.',85,90,80,35,45,30,25,55,35,55,90,65,25,35,'NUM<45:25'],
 ['finance','Keuangan','Bisnis & Ekonomi',true,60,false,'Mengelola investasi, risiko, dan keputusan modal.','Kompetitif dan intensif angka.',90,95,80,40,50,30,30,65,40,65,100,55,25,40,'NUM<50:30'],
 ['business-digital','Bisnis Digital','Bisnis & Ekonomi',true,56,true,'Membangun model bisnis berbasis teknologi.','Perlu adaptif pada tren dan eksperimen pasar.',70,55,55,35,75,35,70,85,70,85,100,35,25,40,'BUS<35:20'],
 ['entrepreneurship','Kewirausahaan','Bisnis & Ekonomi',true,54,true,'Menciptakan dan mengembangkan usaha.','Risiko tinggi; belajar paling kuat lewat eksekusi nyata.',55,45,45,25,45,35,80,90,80,95,100,40,25,45,''],
 ['law','Ilmu Hukum','Hukum & Kebijakan',true,60,false,'Menganalisis aturan, argumen, dan keadilan.','Banyak membaca, menulis, dan berargumentasi.',75,35,85,30,25,20,50,95,80,80,65,100,35,55,'LAW<50:30;COM<40:20'],
 ['international-relations','Hubungan Internasional','Hukum & Kebijakan',true,58,false,'Mempelajari diplomasi, konflik, dan ekonomi global.','Perlu membaca luas, bahasa, dan komunikasi kuat.',70,35,80,35,25,20,55,95,90,75,70,85,35,65,'COM<40:20'],
 ['public-policy','Kebijakan Publik/Administrasi Publik','Hukum & Kebijakan',true,56,true,'Merancang dan mengevaluasi kebijakan pemerintah.','Memadukan data, politik, hukum, dan implementasi.',80,55,90,40,35,25,45,85,95,85,70,90,45,75,''],
 ['criminology','Kriminologi','Hukum & Kebijakan',true,57,true,'Mengkaji kejahatan, kontrol sosial, dan keadilan.','Berhadapan dengan isu sensitif dan riset sosial.',80,45,90,45,35,25,45,80,90,65,50,90,45,55,''],
 ['communication','Ilmu Komunikasi','Media & Komunikasi',true,55,false,'Merancang pesan, media, dan hubungan publik.','Portofolio dan pengalaman praktik sangat penting.',45,25,55,25,35,25,80,100,90,70,75,45,30,45,'COM<45:25'],
 ['journalism','Jurnalistik','Media & Komunikasi',true,56,false,'Mengumpulkan fakta dan menceritakan isu publik.','Deadline tinggi; perlu rasa ingin tahu dan verifikasi.',70,30,95,35,35,20,80,100,85,55,50,60,30,55,'COM<45:25;RES<40:20'],
 ['visual-design','Desain Komunikasi Visual','Desain & Kreatif',true,55,false,'Menyampaikan ide melalui visual dan desain.','Portofolio, kritik studio, dan latihan teknis sangat penting.',35,20,45,20,45,35,100,80,55,45,65,25,20,40,'CRT<50:30'],
 ['product-design','Desain Produk','Desain & Kreatif',true,56,true,'Merancang produk yang berguna dan menarik.','Memadukan riset pengguna, estetika, dan prototipe.',60,45,70,40,65,75,100,75,70,60,70,25,25,55,'CRT<45:25'],
 ['film','Film dan Televisi','Desain & Kreatif',true,54,true,'Menciptakan cerita audiovisual melalui kerja tim.','Produksi padat, kolaboratif, dan portofolio-driven.',35,20,45,20,50,40,100,95,80,70,60,25,25,40,'CRT<45:25'],
 ['literature','Sastra/Linguistik','Sosial & Humaniora',true,55,false,'Mengkaji bahasa, teks, budaya, dan makna.','Banyak membaca, menulis, dan analisis kualitatif.',60,20,80,25,20,15,85,95,75,45,40,35,25,45,'COM<40:20'],
 ['education','Pendidikan','Sosial & Humaniora',true,55,false,'Merancang pembelajaran dan mengembangkan peserta didik.','Perlu kesabaran, komunikasi, dan praktik lapangan.',50,35,65,40,35,25,65,95,100,75,45,35,65,50,'SOC<45:25;COM<40:20'],
 ['sociology','Sosiologi','Sosial & Humaniora',true,56,false,'Menganalisis masyarakat, kelompok, dan perubahan sosial.','Banyak teori, riset, dan penulisan.',70,40,90,35,25,20,55,85,100,60,50,65,40,65,'SOC<40:20'],
 ['anthropology','Antropologi','Sosial & Humaniora',true,55,true,'Memahami budaya dan manusia melalui riset mendalam.','Sering membutuhkan observasi lapangan dan refleksi etis.',60,25,95,35,25,20,65,85,100,55,40,50,45,70,'RES<40:20'],
 ['biology','Biologi','Sains',true,60,false,'Mempelajari kehidupan dari sel hingga ekosistem.','Banyak praktikum, riset, dan materi sains.',75,60,95,100,45,40,35,45,45,40,30,25,70,85,'SCI<50:30'],
 ['chemistry','Kimia','Sains',true,60,false,'Mempelajari materi, reaksi, dan analisis laboratorium.','Butuh ketelitian, matematika, dan kerja laboratorium.',80,75,95,100,50,55,25,35,30,35,30,25,55,65,'SCI<50:30'],
 ['physics','Fisika','Sains',true,62,false,'Menjelaskan fenomena alam dengan model matematis.','Sangat intensif matematika dan abstraksi.',95,100,95,100,70,75,35,35,25,40,35,20,25,60,'NUM<50:30;SCI<50:30'],
 ['biotechnology','Bioteknologi','Sains',true,61,true,'Menggunakan sistem biologis untuk pangan, kesehatan, dan industri.','Riset laboratorium dan sains dasar cukup kuat.',80,65,100,100,65,60,40,45,40,45,45,30,80,75,'SCI<50:30'],
 ['agriculture','Agribisnis/Agroteknologi','Lingkungan',true,56,true,'Mengembangkan pangan, budidaya, dan bisnis pertanian.','Memadukan sains, lapangan, bisnis, dan ketahanan lingkungan.',65,55,75,80,50,55,45,65,70,70,80,40,55,100,'ENV<35:20'],
 ['forestry','Kehutanan','Lingkungan',true,56,true,'Mengelola hutan, konservasi, dan sumber daya alam.','Banyak kerja lapangan dan isu kebijakan lingkungan.',60,45,75,80,40,45,40,65,75,65,50,55,50,100,'ENV<45:25'],
 ['marine-science','Ilmu Kelautan','Lingkungan',true,58,true,'Mempelajari laut, pesisir, dan sumber daya perairan.','Butuh sains dan kesiapan kegiatan lapangan.',65,55,85,95,50,55,40,55,60,55,45,35,50,100,'ENV<45:25;SCI<40:20'],
 ['geography','Geografi','Lingkungan',true,56,true,'Mengkaji ruang, manusia, lingkungan, dan data geospasial.','Gabungan lapangan, analisis spasial, dan isu sosial.',75,60,85,65,60,45,55,70,75,60,55,55,35,100,'ENV<35:20']
];
