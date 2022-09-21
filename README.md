# Sportident za priznavanje KTjev na taborniških orientacijah

Ta repozitorij je namenjen ljudem, ki si želijo olajšati življenje nočejo na roke preverjati, koliko KT so ekipe uspešno
našle. Za uporabo tega programa rabiš:
- Bazno postajo za branje čipov,
- Programsko opremo SPORTIdent Config+,
- Python3.

Program za svoje delovanje potrebuje naslednje knjižnice:
- **numpy** (delo s tabelami),
- **openpyxl** (pisanje v xlsx datoteko),
- **os** ter **shutil** (kopiranje 'readcard' tabele).

# Struktura datotek
Datoteke, ki so povezane z neko tekmo so zbrane v eni mapi (v tem primeru je to 
test_system). V tej mapi so:
- log datoteke posameznih ekip v `logs/_stevilka ekipe_.txt` ter združen log `logger.txt`,
- Seznam ekip ter SIID številk njihovih čipov v `results_input.xlsx`,
- Trase kategorij v `100.csv, 200.csv, ...`

# Priprava podatkov
1. Prekopiraj mapo `test_system` in spremeni ime
2. Odpri results_input.xlsx in dodaj imena ekip (stolpec A) ter SIID številke (stoplec B). Za zadnjo ekipo more biti 
celica z tekstom 'STOP'.
3. Za vsako kategorijo uredi `X00.csv` datoteko z naslednjimi stoplci:
   - A: koda postaje
   - B: Čas umika kontrolne točke
   - C: Številka KTja
   - D+: dodatne ključne besede (vrstni red ni ključen, vsaka ključna beseda gre v svoj stoplec)
     - 'mrtvi_cas': Na tem KTju se pričakuje, da lahko ekipa prejme mrtvi čas
     - 'hitrostna_start': Start hitrostne etape
     - 'hitrostna_cilj': Cilj hitrostne etape

# Izvoz podatkov v SPORTIdent Config+
Ko se bazna postaja poveže z računalnikom in so preko `read SI-cards` zavikha prebrani vsi čipi moremo izvoziti tabelo z
vsemi podatki. Preko `Export...` &rarr; `Export full detail list`. 

Readcard tabelo lahko shranimo v a) mapo `datadump` ali b) mapo, kjer so ostale datoteke od tekme.

# Poganjanje skripte
Glaven del repozitorija je datoteka `main.py`. V tej datoteki je napisana funkcija `recalculate_results`, ki ustvari
excel datoteko `results_output.xlsx`, s tabelo izračunanih rezultatov. Argumenti funkcije so:
- `folder`: ime mape z vsemi potrebnimi datotekami
- `track_csv_separator`: Separator med celicami v datotekah `X00.csv` (lahko se spremeni zaradi regijskih nastavitev v 
Excelu)
- `automatic_readcard_name`: logičen argument. V primeru, da je enak True, se readcard tabela avtomatsko prekopira iz 
mape `datadump`. Prekopira se najnovejša datoteka `readtable_..._datum_ura.csv.backup`, ki se ustvari, ko se iz programa
Config+ izvozi tabela več kot enkrat (torej je treba tabelo izvoziti večkrat). Če je argument enak False, se tabela
prebere iz `glavna_mapa/readcard_filename`
- `readcard_filename`: Če je `automatic_readcard_name` enak False, potem prebere tabelo iz
`glavna_mapa/readcard_filename`
- `comply_with_deadtime_tag`: Logičen argument. V primeru, da je enak True, se mrtvi čas pri KTjih, kjer ni tag-a 
"mrtvi_cas", ne prišteje. V nasprotnem primeru pa se prišteje. V obeh primerih pa se izpiše opozorilo

**Kako zagnati program?** Ali popravi klic funkcije v zadnji vrstici datoteke `main.py` in jo zaženi z Pythonom ali pa
uvozi datoteko in poženi funkcijo z pravimi argumenti.


# Mrtvi čas
Mrtvi čas je ekipam dodeljen tako, da se ekipa čipira na kontrolni točki dvakrat:
- Ob prihodu ekipe na KT
- V primeru, da ekipi pripada mrtvi čas se pred začetkom opravljanja nalog ekipa čipira še enkrat na isto postajo. 

Mrtvi čas se na dani kontrolni točki šteje le v primeru, da je KT v tabli označen s 
ključno besedo "mrtvi_cas" (in je argument comply_with_deadtime_tag==True). Tako
je preprečeno, da bi si ekipa na mrtvi kontrolni točki dodala mrtvi čas. 

Tudi na živih KTjih, kjer je možno dobiti mrtvi čas, morajo kontrolorji omogočiti le
ekipam, ki "si to zaslužijo".

# Kako deluje program?
Program na podlagi številke ekipe najprej najde kategorijo. Nato odpre datoteko od 
trase in prebere id postaj, maksimalne dovoljene čase ter ključne besede. Nato za vsak KT preveri
mrtvi čas in na podlagi časa starta ekipe izračuna, če so KTji bili znotraj maksimalnega dovoljenega časa
ali ne.
