import pandas as pd
import numpy as np
def hakem_atamasi():
    epak_listesi = "EPAK listesi.xlsx"
    hakem_mazeret_listesi = "hakem mazeret listesi.xlsx"
    antrenmana_katılanlar_listesi = "antrenmana katılanlar listesi.xlsx"
    antrenman_günü_maçı_olanlar_listesi = "antrenman günü maçı olanlar listesi.xlsx"

    epak_listesi_oku = pd.read_excel(epak_listesi)
    hakem_mazeret_listesi_oku = pd.read_excel(hakem_mazeret_listesi)
    antrenmana_katılanlar_listesi_oku = pd.read_excel(antrenmana_katılanlar_listesi)
    antrenman_günü_maçı_olanlar_listesi_oku = pd.read_excel(antrenman_günü_maçı_olanlar_listesi)

    maç_atanacak_hakemler = []

    #EPAK'a katılıp katılmadığını kontrol ederek katılanların bir listesini oluştur
    epak_katılan_hakemler = []
    for index, satir in epak_listesi_oku.iterrows():
        if isinstance(satir.iloc[5], (str,float)) and str(satir.iloc[5]).upper() == "X":
            epak_katılan_hakemler.append(satir.iloc[1])


    tarih = input("Mazeret kontrol edilecek tarihi girin(gün.ay.yıl olarak): ")
    #girilen tarihte hakemin mazereti var mı diye kontrol etmek için
    tarihler = hakem_mazeret_listesi_oku.iloc[1::4, 2].tolist() #dosyanın 3. sütununda 3.satırdan itibaren her 4 satırda bir mazeret tarihleri yazıyor
    isimler = hakem_mazeret_listesi_oku.iloc[1::4, 0].tolist() #dosyanın 1. sütununda 3.satırdan itibaren her 4 satırda bir isimler yazıyor

    #girilen tarihte mazereti olan hakemleri listeden çıkarmak için
    #silinecek hakemler adlı geçici bir liste açıp girilen tarihte
    #mazereti olan hakemleri bu listeye koyma
    silinecek_hakemler = []
    for hakem, mazeret_tarihi in zip(isimler, tarihler):
                if hakem in maç_atanacak_hakemler and mazeret_tarihi == tarih:
                    silinecek_hakemler.append(hakem)

    #antrenman günü maça çıkanları al ve tekrarlanan isimleri tek bir değere düşür
    antrenman_günü_maçı_olanlar = []
    antrenman_günü_maçı_olanlar_isimleri = antrenman_günü_maçı_olanlar_listesi_oku.iloc[2:, [7, 8, 9]].values.tolist()

    for hakemler in antrenman_günü_maçı_olanlar_isimleri:
        for hakem in hakemler:
            if isinstance(hakem ,str):
                antrenman_günü_maçı_olanlar.append(hakem)

    unique_set = set(antrenman_günü_maçı_olanlar)
    antrenman_günü_maçı_olanlar_unique = list(unique_set)

    #antrenmana katılanları al
    antrenmana_katılanlar = antrenmana_katılanlar_listesi_oku.iloc[:, 0].tolist()

    #bütün şartları kontrol ederek girilen tarihte maç atanması uygun olan hakemleri listele
    for hakem in epak_katılan_hakemler:
        if hakem in silinecek_hakemler:
            pass
        elif hakem not in silinecek_hakemler:
            if hakem in antrenman_günü_maçı_olanlar_unique or hakem in antrenmana_katılanlar:
                maç_atanacak_hakemler.append(hakem)

    #maç atanacak hakemler listesini excele dönüştür
    df = pd.DataFrame({"Hakemler": maç_atanacak_hakemler})
    df.to_excel("maç atanacak hakemler.xlsx", index=False, header=False)

    return maç_atanacak_hakemler

def mac_atamasi():
    hakem_listesi = []  # Önce tanımla
    hakem_listesi = hakem_atamasi()  # Hakem listesini baştan al

    mac_listesi_exceli = "maçlar.xlsx"
    mac_listesi_exceli_oku = pd.read_excel(mac_listesi_exceli)
    mac_listesi_secilen_sutunlar = mac_listesi_exceli_oku.iloc[27:, [2, 3, 4, 5, 6]]
    mac_listesi = mac_listesi_secilen_sutunlar.values.tolist()

    atanacak_hakemler = []
    hakemli_mac_listesi = []
    kalan_hakemler = []  # Atanamayan (kalan) hakemlerin listesi
    previous_field = None  # To keep track of the previous field

    for maclar in mac_listesi:
        ikinci_eleman = maclar[2]
        current_field = maclar[0]  # Assuming the field is in the 4th column

        if ikinci_eleman in ['U 13 ELİT', 'U 13 1.KÜME']:
            if not hakem_listesi:
                print("Warning: No more referees available. Resetting referee list.")
                hakem_listesi = hakem_atamasi()  # Hakem listesini baştan al

            # If the field is the same as the previous match, use the same referee
            if current_field == previous_field:
                hakem = atanacak_hakemler[-1]
            else:
                hakem = hakem_listesi.pop(0)
                atanacak_hakemler.append(hakem)

            # Append the match along with the assigned referee to hakemli_mac_listesi
            maclar_with_hakem = maclar + [hakem]
            hakemli_mac_listesi.append(maclar_with_hakem)

            # Update the previous field for the next iteration
            previous_field = current_field
        else:
            print("Skipping non-'U 13 ELİT' match")

    # Atanamayan (kalan) hakemleri ekle
    kalan_hakemler.extend(hakem_listesi)

    # Hakemli maçları yazdır
    for maclar in hakemli_mac_listesi:
        print(maclar)

    # Atanamayan (kalan) hakemleri yazdır
    print("Kalan Hakemler:")
    for hakem in kalan_hakemler:
        print(hakem)


mac_atamasi()


