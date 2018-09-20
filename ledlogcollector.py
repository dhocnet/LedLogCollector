#!
# Led Log Collector
# Versi: 0.0.1-alpha
# Oleh: Heru Niagara
# Deskripsi:
#     - Untuk mengumpulkan playlog LM 2012
#     - Hasil di eksport ke xml sehingga bisa di import
#       ke LibreOffice Calc atau MS. Office Excel
#
# Info: https://www.dhocnet.work
# Kontak: mongkee.lutfi@gmail.com
#

import os, codecs, base64

# variable OS 32 dan 64-bit
llc_32 = "c:\\Program Files\\LED soft\\LED Manager 2012\\LogFile"
llc_64 = "c:\\Program Files (x86)\\LED soft\\LED Manager 2012\\LogFile"
llc_out = "c:\\LedLog Collector"
llc_xmlh = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><PlayLog xmlns:xsi=\"http://www.w3.org/2001/XMLSchema\">"

def llc_kerja(llc_ar):
	llc_fdir = os.listdir(llc_ar)
	print "Ditemukan: %d direktori log..."%len(llc_fdir)
	for llc_pil in llc_fdir:
        print "Memeriksa eksistensi direktori hasil..."
		if not os.path.exists("%s\\%s"%(llc_out,llc_pil)):
            print "Direktori %s tidak ditemukan!"%llc_pil
            print "Membuat direktori %s untuk menyimpan hasil..."%llc_pil,
			os.mkdir("%s\\%s"%(llc_out,llc_pil))
			print "OK"
		llc_ffile = os.listdir("%s\\%s"%(llc_ar,llc_pil))
		for llc_ffx in llc_ffile:
			if os.path.exists("%s\\%s\\%sxml"%(llc_out,llc_pil,llc_ffx[:-3])):
				print "File log: %s telah diproses. Melewati,..." %llc_ffx
			else:
                print "Mengurai file log: %s"%llc_ffx
				llc_a = codecs.open("%s\\%s\\%s"%(llc_ar,llc_pil,llc_ffx),"r",encoding="utf-16")
				llc_b = llc_a.readlines()
				llc_a.close()
				print "Ditemukan %d baris untuk diproses..."%len(llc_b)
				llc_a = codecs.open("%s\\%s\\%s.xml"%(llc_out,llc_pil,llc_ffx[:-4]),"w",encoding="utf-16")
				llc_a.writelines(llc_xmlh)
				for llc_c in llc_b:
					if llc_c.find("Duration") != -1:
						llc_e = llc_c.split("=")
						llc_f = llc_e[0][:8]
						llc_g = llc_e[1][:-7]
						llc_a.write("<PlayData><WaktuMain>%s</WaktuMain><FileMedia>%s</FileMedia></PlayData>"%(llc_f,llc_g))
						print "Ditambahkan: %s"%llc_g
					else:
						pass
                llc_a.writelines("</PlayLog>")
				llc_a.close()
            print "%s telah selesai diproses."%llc_ffx
        print "Log file tahun %s telah diproses semua."%llc_pil[-4:]
    print "Pekerjaan telah selesai!"
    llc_bantuan()

def llc_bantuan():
    print "\n\nSELAMAT! PEKERJAAN TELAH SELESAI!\n\nBerkas hasil kerja bisa Anda temukan di c:/LedLog Collector. Berkas disimpan dalam format XML. Anda dapat membukanya dengan program LibreOffice Calc, Google Spreadsheet dan MS. Office Excel. Caranya, buka terlebih dahulu program office Anda lalu pilih menu [File] -> [Import XML Documents].\n\nOleh: mongkeelutfi\nKontribusi: https://github.com/dhocnet/ledlogcollector\nWebsite: https://www.dhocnet.work\nEmail: mongkee.lutfi@gmail.com\n"

if os.path.exists(llc_32):
    llc_kerja(llc_32)
elif os.path.exists(llc_64):
    llc_kerja(llc_64)
else:
    print "LED Manager tidak ditemukan!"
