//pozn.: mazání prvního obrázku v dokumentu: SpreadsheetApp.getActiveSheet().getImages()[0].remove();

//Testování proměnných:
//Browser.msgBox(promenna, Browser.Buttons.OK_CANCEL);
//return;

//nastavení cache (jednoduchá náhrada global variables)
var cache = CacheService.getPrivateCache(); 

//-------------------------- PŘIDÁNÍ MENU ---------------------------------

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{name : "Vygeneruj řádek rozvrhu", functionName : "Pridel_radek"},
                 {name : "Vytiskni rozvrh", functionName : "Priprav_rozvrh_na_tisk"},
                 {name : "Připrav předsedání ŽaS", functionName : "Priprav_zas_na_tisk"}];
                 
  sheet.addMenu("Koordinátor", entries);
};
//{name : "Importuj program ŽaS", functionName : "Importuj_program"}]; (a i ten hárok je skrytý)

//-------------------------- FUNKCE ---------------------------------

function getElementsByID(element, idToFind) {  
  var regId = new RegExp( '(<[^<]*id=[\'"]'+ idToFind +'[\'"][^>]*)' );
  var result = regId.exec( element );
  return result[1] + '>';
}

function Pondelky(year, month) {

    var day, counter, date;

    day = 1;
    counter = 0;
    date = new Date(year, month, day);
    while (date.getMonth() === month) {
        if (date.getDay() === 1) { // Sun=0, Mon=1, Tue=2, etc.
            counter += 1;
        }
        day += 1;
        date = new Date(year, month, day);
    }
    return counter;
}

function Datumy_pondelku (year, month) {

    var day, counter, date;
    
    var vysledek = [];

    day = 1;
    counter = 0;
    date = new Date(year, month, day);
    while (date.getMonth() === month) {
        if (date.getDay() === 1) { // Sun=0, Mon=1, Tue=2, etc.
            counter += 1;
            vysledek.push(date);
        }
        day += 1;
        date = new Date(year, month, day);
    }
    return vysledek;
}

function Kolikaty_tyden (year, month, datumko) {

    var day, counter, date, vysledek;
    
    vysledek = 0;
    day = 1;
    counter = 0;
    date = new Date(year, month, day);
    while (date.getMonth() === month) {
        if (date.getDay() === 1) { // Sun=0, Mon=1, Tue=2, etc.
            counter += 1;
           
            if (date.getTime() == datumko.getTime()) {
              vysledek = counter;
            }    
          
        }
        day += 1;
        date = new Date(year, month, day);
    }
    return vysledek;
}

//---------------------- IMPORTOVÁNÍ PROGRAMU ŽAS PRO PJ ------------------------------------------------------------------------------------------------------------------------------
function Importuj_program() {
  
    var spreadsheet = SpreadsheetApp.getActive();
  
    // importování, úprava a nasypání do proměnné "vysledek"
    var posledni_radek_import = spreadsheet.getSheetByName('Import').getDataRange().getValues().length;
    var sloupec_data = spreadsheet.getSheetByName('Import').getRange('A1:A' + posledni_radek_import).getValues();
  
    var vysledek = [];
    var retezec = "";
    var i;
    var n;
   
    for(n = 0; n < sloupec_data.length; n++){ 
     
     if(sloupec_data[n+1][0] != "" && sloupec_data[n][0] != "") {
      //pozn. pokud je splněna podmínka u for, provádí se to. Když podmínka neplatí, už se to neprovede.       
       retezec = "";
       for (i = n; sloupec_data[i][0] != ""; i++) { 
         if (retezec == "") {
           retezec = sloupec_data[i][0];
         } else { 
           retezec = retezec + " " + sloupec_data[i][0];
         }
         if (i == sloupec_data.length-1){
           break;
         }
       }     
       vysledek.push(retezec);
       n = i;
     }     
    }
  
    // nalití do tabulky
    var posledni_radek_program = spreadsheet.getSheetByName('(Program ŽaS)').getDataRange().getValues().length;
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('(Program ŽaS)'), true);
    var prvni_radek_novy = posledni_radek_program + 1;
  
    var pocitadlo = 0; 
    for (n=0; n < vysledek.length/6; n++){                                                        //2019 - ze 4 urobeno 6 (protože je 6 sloupců)
      for (i = 1; i < 7; i++) {                                                                   //2019 z 5 urobeno 7
        spreadsheet.getRange("R" + prvni_radek_novy + "C" + i).setValue(vysledek[pocitadlo]);
        pocitadlo++;
      } 
      prvni_radek_novy = prvni_radek_novy + 1;
    }

  
}  

//---------------------- PŘÍPRAVA PŘEDSEDÁNÍ ŽAS NA TISK ---------------

function Priprav_zas_na_tisk() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  //kontrola, jestli nestojíme kurzorem mimo sloupec A
  if (spreadsheet.getActiveCell().getColumn()>1) {
     Browser.msgBox("Nestojíš v sloupci A, končím.", Browser.Buttons.OK_CANCEL);
     return;
  }
  
  //zjištění datumu
  var datum = spreadsheet.getCurrentCell().getValue();
  var pondelni_datum = spreadsheet.getCurrentCell().getValue();  
  var mesic = datum.getMonth()+1;
  var rok = datum.getFullYear();
  var den = datum.getDate();
  
  //který týden je to v měsíci? (počítáno od prvního pondělku) (přidáno 2019)
  var ktery_tyden = Kolikaty_tyden(rok, mesic-1,datum); //0=leden, 11=prosinec, proto odečítáme 1  
 
  //fetchnutí WOL s rozvrhem na daný den
  var url = 'https://wol.jw.org/sk/wol/dt/r38/lp-v/' + rok + '/' + mesic + '/' + den;
  
  var html = UrlFetchApp.fetch(url).getContentText();  
  
  //zparsování obsahů programu jedna a dva

  //odříznutí celé patičky
  var ocistene_html = html;
  var pozice_paticky = ocistene_html.indexOf('<div id="regionFooter');
  ocistene_html = ocistene_html.slice(0, pozice_paticky);
  //odříznutí celého úvodu
  var pozice_uvodu = ocistene_html.indexOf('<div id="regionMain');
  ocistene_html = ocistene_html.slice(pozice_uvodu);  
  
  //nalezení pozic bodů ul s class "noMarker"
  var vyskyty_nomarker = [];
  for (i = 0; i < ocistene_html.length; ++i) {
    if (ocistene_html.substring(i, i + 'noMarker'.length) == 'noMarker') {
      vyskyty_nomarker.push(i-11);
    }
  }  
  //počet ul s class noMarker
  var pocet_nomarker = vyskyty_nomarker.length;
  
  //1) vyhodíme všechny tyto uly s class noMarker
  //nalezení všech špatných ulů
  var spatny_ul = [];
  for (i = 0; i < pocet_nomarker; ++i) {
    //odříznutí začátku stringu
    var zbytek_ul = ocistene_html.slice(vyskyty_nomarker[i]);
    //nalezení nejbližšího uzavíracího tagu ul
    var kde_je_uzaviraci_ul = zbytek_ul.indexOf("</ul>");
    //vystřihnutí obsahu ulu
    spatny_ul.push(zbytek_ul.slice(0, kde_je_uzaviraci_ul+5));
  } 
  //vyříznutí všech špatných ulů
  for (i = 0; i < pocet_nomarker; ++i) {
    ocistene_html = ocistene_html.replace(spatny_ul[i],"");
  } 
  
  //2) hledáme <li>, které už představují jen opuntíkované body
  var vyskyty_li = [];
  for (i = 0; i < ocistene_html.length; ++i) {
    if (ocistene_html.substring(i, i + '<li>'.length) == '<li>') {
      vyskyty_li.push(i);
    }
  } 
    
  //2B) hledáme <header> ve kterém je co se čte z Bible (to je v h2)
  var vyskyty_header = [];
  for (i = 0; i < ocistene_html.length; ++i) {
    if (ocistene_html.substring(i, i + '<header>'.length) == '<header>') {
      vyskyty_header.push(i);
    }
  }
  //obsah druhého headeru
  var obsah_header = ocistene_html.slice(vyskyty_header[1]);
  var kde_je_uzaviraci = obsah_header.indexOf("</header>");
  obsah_header = obsah_header.slice(0, kde_je_uzaviraci+9);
  //obsah h2
  var obsah_h2 = "<h2" + obsah_header.split("<h2")[1];
  var tydenni_cteni=obsah_h2.replace(/<[^>]*>/g, '');
  
  //3) pokud je vyskyty_li.length=13, je jen jeden program, pokud 14, jsou dva programy - pozn. 2019: to platí třetí a další týden. Neplatí to 1. a 2. týden. První týden platí: 12/13 a druhý týden platí 14/15.
  //-------------- program 1 nebo i 2? -----------------
  //rok 2019 - oprava podle toho, který je to týden
  var oprava2019 = 0;
  
  if (ktery_tyden == 1) {
    oprava2019 = -1;
  }
  
  if (ktery_tyden == 2) {
    oprava2019 = 1;
  } 
  
  if (vyskyty_li.length == (14+oprava2019)){
    var dvaprogramy = true;
  } else {
    var dvaprogramy = false;
  }
  //vysledek: dvaprogramy (pokud je true, tak jsou dva programy)
  
  //4)získáme obsah bodů programu (body jsou číslované od nuly)
  var obsah_prejav = ocistene_html.slice(vyskyty_li[2]);
  var kde_je_uzaviraci = obsah_prejav.indexOf("</li>");
  obsah_prejav = obsah_prejav.slice(0, kde_je_uzaviraci+5);   
  //převod HTML na plaintext
  obsah_prejav=obsah_prejav.replace(/<[^>]*>/g, '');
  
  //v roce 2019 už čtení Bible od slyšících nepotřebujeme
  //var obsah_ctenibi = ocistene_html.slice(vyskyty_li[4]);
  //var kde_je_uzaviraci = obsah_ctenibi.indexOf("</li>");
  //obsah_ctenibi = obsah_ctenibi.slice(0, kde_je_uzaviraci+5);   
  //převod HTML na plaintext
  //obsah_ctenibi=obsah_ctenibi.replace(/<[^>]*>/g, ''); 
  
  //získáme obsah programů 1 a 2
  var obsah_program1 = ocistene_html.slice(vyskyty_li[9+oprava2019]);
  var kde_je_uzaviraci = obsah_program1.indexOf("</li>");
  obsah_program1 = obsah_program1.slice(0, kde_je_uzaviraci+5);   
 
  //převod HTML na plaintext
  obsah_program1=obsah_program1.replace(/<[^>]*>/g, '');
  
  if (dvaprogramy == true) {
    var obsah_program2 = ocistene_html.slice(vyskyty_li[10+oprava2019]);
    var kde_je_uzaviraci = obsah_program2.indexOf("</li>");
    obsah_program2 = obsah_program2.slice(0, kde_je_uzaviraci+5);   
    //převod HTML na plaintext
    obsah_program2=obsah_program2.replace(/<[^>]*>/g, '');
    
    var obsah_stubi = ocistene_html.slice(vyskyty_li[11+oprava2019]);
    var kde_je_uzaviraci = obsah_stubi.indexOf("</li>");
    obsah_stubi = obsah_stubi.slice(0, kde_je_uzaviraci+5);   
    //převod HTML na plaintext
    obsah_stubi=obsah_stubi.replace(/<[^>]*>/g, '');
  } else {
    var obsah_program2 = "-";
    var obsah_stubi = ocistene_html.slice(vyskyty_li[10+oprava2019]);
    var kde_je_uzaviraci = obsah_stubi.indexOf("</li>");
    obsah_stubi = obsah_stubi.slice(0, kde_je_uzaviraci+5);   
    //převod HTML na plaintext
    obsah_stubi=obsah_stubi.replace(/<[^>]*>/g, '');      
  }  
   
  //range s rozvrhem do array
  //nalezení posledního sloupce rozvrhu
  var posledni_sloupec_rozvrh = spreadsheet.getLastColumn();
  //nalezení současného řádku
  var soucasny_radek_rozvrh = spreadsheet.getActiveCell().getRow();
  var dalsi_radek = soucasny_radek_rozvrh + 1; 
  var array_rozvrh = spreadsheet.getRange('R' + soucasny_radek_rozvrh + 'C1:R' + dalsi_radek + 'C' + posledni_sloupec_rozvrh).getValues();
  
  //přečtení nastavení který den je shromáždění přes týden
  var den_tyden = spreadsheet.getSheetByName("Nastavení").getRange('B3').getValues();
  
  //nalezení datumu týdenního 
  var datum_tydenniho = datum;
  datum_tydenniho.setDate(datum_tydenniho.getDate() + Number(den_tyden));
  
  //zapsání datumu do titulku
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('(Tisk přípravy ŽaS)'), true);
  spreadsheet.getRange("C1").setValue(datum_tydenniho);
  
  //doplnění informací k písni 1
  var posledni_radek_pisne = spreadsheet.getSheetByName('(Písně sny55)').getDataRange().getValues().length;
  var sloupec_cisla_pisni = spreadsheet.getSheetByName('(Písně sny55)').getRange('A2:A' + posledni_radek_pisne).getValues();
  //číslo řádku v písních, na kterém je stejné číslo
  for(var n in sloupec_cisla_pisni){
   if(sloupec_cisla_pisni[n][0] == array_rozvrh[0][1]) {
    var radek_v_pisnich = 2 + Number(n);
    break;
   }
  }  
  //přečtení názvu, verše a v PJ
  var nazev = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 2).getValues();
  nazev = nazev.toString(); //převod z array na string
  var vers = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 3).getValues();
  vers = vers.toString(); //převod z array na string
  var v_pj = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 4).getValues();
  v_pj = v_pj.toString(); //převod z array na string
  //sestavení položky do rozvrhu
  array_rozvrh[0][1] = array_rozvrh[0][1] + " - " + nazev + " (" + vers + " " + v_pj + ")";
 
  var posun_rozvrh_2019 = 3; //protože v rozvrhu 2019 se objevuje navíc úloha 4, která zabírá 3 řídky navíc
  
  //doplnění informací k písni 2
  for(var n in sloupec_cisla_pisni){
   if(sloupec_cisla_pisni[n][0] == array_rozvrh[0][17+posun_rozvrh_2019]) {
    var radek_v_pisnich = 2 + Number(n);
    break;
   }
  }  
  //přečtení názvu, verše a v PJ
  var nazev = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 2).getValues();
  nazev = nazev.toString(); //převod z array na string
  var vers = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 3).getValues();
  vers = vers.toString(); //převod z array na string
  var v_pj = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 4).getValues();
  v_pj = v_pj.toString(); //převod z array na string
  //sestavení položky do rozvrhu
  array_rozvrh[0][17+posun_rozvrh_2019] = array_rozvrh[0][17+posun_rozvrh_2019] + " - " + nazev + " (" + vers + " " + v_pj + ")";  
  
  //doplnění informací k písni 3
  for(var n in sloupec_cisla_pisni){
   if(sloupec_cisla_pisni[n][0] == array_rozvrh[0][22+posun_rozvrh_2019]) {
    var radek_v_pisnich = 2 + Number(n);
    break;
   }
  }  
  //přečtení názvu, verše a v PJ
  var nazev = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 2).getValues();
  nazev = nazev.toString(); //převod z array na string
  var vers = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 3).getValues();
  vers = vers.toString(); //převod z array na string
  var v_pj = spreadsheet.getSheetByName('(Písně sny55)').getRange(radek_v_pisnich, 4).getValues();
  v_pj = v_pj.toString(); //převod z array na string
  //sestavení položky do rozvrhu
  array_rozvrh[0][22+posun_rozvrh_2019] = array_rozvrh[0][22+posun_rozvrh_2019] + " - " + nazev + " (" + vers + " " + v_pj + ")"; 

  //doplnění informací k znaku čtenáře
  var posledni_radek_znaky = spreadsheet.getSheetByName('(Znaky)').getDataRange().getValues().length;
  var sloupec_cisla_znaku = spreadsheet.getSheetByName('(Znaky)').getRange('A2:A' + posledni_radek_znaky).getValues();
  //číslo řádku v znacích, na kterém je stejné číslo
  for(var n in sloupec_cisla_znaku){
   if(sloupec_cisla_znaku[n][0] == array_rozvrh[0][7]) {
    var radek_v_znacich = 2 + Number(n);
    break;
   }
  }  
  //přečtení názvu, popisu
  if (array_rozvrh[0][7] == "-" || array_rozvrh[0][7] == ""){
    } else { 
    var nazev_znaku = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 2).getValues();
    nazev_znaku = nazev_znaku.toString(); //převod z array na string
    var popis = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 3).getValues();
    popis = popis.toString(); //převod z array na string
    //sestavení položky do rozvrhu
    array_rozvrh[0][7] = array_rozvrh[0][7] + " - " + nazev_znaku + " (=" + popis + ")"; 
  }
  
  //doplnění informací k znaku 1
  if (array_rozvrh[0][9] == "-" || array_rozvrh[0][9] == ""){
  } else { 
    for(var n in sloupec_cisla_znaku){
      if(sloupec_cisla_znaku[n][0] == array_rozvrh[0][9]) {
        var radek_v_znacich = 2 + Number(n);
        break;
      }
    }  
    //přečtení názvu, popisu
    var nazev_znaku = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 2).getValues();
    nazev_znaku = nazev_znaku.toString(); //převod z array na string
    var popis = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 3).getValues();
    popis = popis.toString(); //převod z array na string
    //sestavení položky do rozvrhu
    array_rozvrh[0][9] = array_rozvrh[0][9] + " - " + nazev_znaku + " (=" + popis + ")";  
  }  
  
  //doplnění informací k znaku 2
  if (array_rozvrh[0][12] == "-" || array_rozvrh[0][12] == ""){
  } else { 
    for(var n in sloupec_cisla_znaku){
     if(sloupec_cisla_znaku[n][0] == array_rozvrh[0][12]) {
      var radek_v_znacich = 2 + Number(n);
      break;
     }
    }  
    //přečtení názvu, popisu
    var nazev_znaku = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 2).getValues();
    nazev_znaku = nazev_znaku.toString(); //převod z array na string
    var popis = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 3).getValues();
    popis = popis.toString(); //převod z array na string
    //sestavení položky do rozvrhu
    array_rozvrh[0][12] = array_rozvrh[0][12] + " - " + nazev_znaku + " (=" + popis + ")";  
  }
  
  //doplnění informací k znaku 3
  if (array_rozvrh[0][15] == "-" || array_rozvrh[0][15] == ""){
  } else { 
    for(var n in sloupec_cisla_znaku){
     if(sloupec_cisla_znaku[n][0] == array_rozvrh[0][15]) {
      var radek_v_znacich = 2 + Number(n);
      break;
     }
    }  
    //přečtení názvu, popisu
    var nazev_znaku = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 2).getValues();
    nazev_znaku = nazev_znaku.toString(); //převod z array na string
    var popis = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 3).getValues();
    popis = popis.toString(); //převod z array na string
    //sestavení položky do rozvrhu
    array_rozvrh[0][15] = array_rozvrh[0][15] + " - " + nazev_znaku + " (=" + popis + ")";
  }
  
   //doplnění informací k znaku 4 (přidáno 2019)
  if (array_rozvrh[0][18] == "-" || array_rozvrh[0][18] == ""){
  } else { 
    for(var n in sloupec_cisla_znaku){
     if(sloupec_cisla_znaku[n][0] == array_rozvrh[0][18]) {
      var radek_v_znacich = 2 + Number(n);
      break;
     }
    }  
    //přečtení názvu, popisu
    var nazev_znaku = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 2).getValues();
    nazev_znaku = nazev_znaku.toString(); //převod z array na string
    var popis = spreadsheet.getSheetByName('(Znaky)').getRange(radek_v_znacich, 3).getValues();
    popis = popis.toString(); //převod z array na string
    //sestavení položky do rozvrhu
    array_rozvrh[0][18] = array_rozvrh[0][18] + " - " + nazev_znaku + " (=" + popis + ")";
  }
    
  //vytáhnutí programu ŽaS pro PJ
  var posledni_radek_program_zas = spreadsheet.getSheetByName('(Program ŽaS)').getDataRange().getValues().length;
  var sloupec_data = spreadsheet.getSheetByName('(Program ŽaS)').getRange('A2:A' + posledni_radek_program_zas).getValues();
  //číslo řádku v programu, na kterém je stejné datum
  for(var n in sloupec_data){
   if(sloupec_data[n][0].getTime() == pondelni_datum.getTime()) {
    var radek_v_programu_zas = 2 + Number(n);
    break;
   }
  }  
  //pokud není doplněn rozvrh ŽaS
  if (radek_v_programu_zas === undefined) {
    Browser.msgBox("Nie je doplnený program ŽaS pre posunkový jazyk!", Browser.Buttons.OK_CANCEL);
    return;
  }
  //přečtení názvu, verše a v PJ (2019 upraveno)
  var obsah_ctenibi = spreadsheet.getSheetByName('(Program ŽaS)').getRange(radek_v_programu_zas, 2).getValues();
  obsah_ctenibi = obsah_ctenibi.toString(); //převod z array na string
  var uloha1 = spreadsheet.getSheetByName('(Program ŽaS)').getRange(radek_v_programu_zas, 3).getValues();
  uloha1 = uloha1.toString(); //převod z array na string
  var uloha2 = spreadsheet.getSheetByName('(Program ŽaS)').getRange(radek_v_programu_zas, 4).getValues();
  uloha2 = uloha2.toString(); //převod z array na string
  var uloha3 = spreadsheet.getSheetByName('(Program ŽaS)').getRange(radek_v_programu_zas, 5).getValues();
  uloha3 = uloha3.toString(); //převod z array na string
  var uloha4 = spreadsheet.getSheetByName('(Program ŽaS)').getRange(radek_v_programu_zas, 6).getValues();
  uloha4 = uloha4.toString(); //převod z array na string
 
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('(Tisk přípravy ŽaS)'), true);
   
  //vypsání arraye s rozvrhem
  spreadsheet.getRange('R2C2').setValue(array_rozvrh[0][1]);//píseň 1
  spreadsheet.getRange('R3C2').setValue(array_rozvrh[0][2]);
  spreadsheet.getRange('R4C2').setValue(array_rozvrh[0][3]);
  
  spreadsheet.getRange('R6C2').setValue(array_rozvrh[0][4]);
  spreadsheet.getRange('R7C2').setValue(array_rozvrh[0][5]);
  spreadsheet.getRange('R8C2').setValue(array_rozvrh[0][6]);
  spreadsheet.getRange('R8C3').setValue(array_rozvrh[0][7]);//znak čtenář
  
  spreadsheet.getRange('R10C2').setValue(array_rozvrh[0][8]);  
  spreadsheet.getRange('R10C3').setValue(array_rozvrh[0][9]);//znak 1
  spreadsheet.getRange('R11C2').setValue(array_rozvrh[0][10]);
  
  spreadsheet.getRange('R12C2').setValue(array_rozvrh[0][11]);
  spreadsheet.getRange('R12C3').setValue(array_rozvrh[0][12]);//znak 2  
  spreadsheet.getRange('R13C2').setValue(array_rozvrh[0][13]);
  
  spreadsheet.getRange('R14C2').setValue(array_rozvrh[0][14]); //úloha 3
  spreadsheet.getRange('R14C3').setValue(array_rozvrh[0][15]); //znak 3
  spreadsheet.getRange('R15C2').setValue(array_rozvrh[0][16]); //partner 3
  
  spreadsheet.getRange('R16C2').setValue(array_rozvrh[0][17]); //úloha 4 - přidáno 2019
  spreadsheet.getRange('R16C3').setValue(array_rozvrh[0][18]); //znak 4 - přidáno 2019
  spreadsheet.getRange('R17C2').setValue(array_rozvrh[0][19]); //partner 4 - přidáno 2019
  
  //dále všechny R+2 protože přibyly 2 řádky kvůli úloze 4
  
  spreadsheet.getRange('R19C2').setValue(array_rozvrh[0][17+posun_rozvrh_2019]);//píseň 2
  spreadsheet.getRange('R20C2').setValue(array_rozvrh[0][18+posun_rozvrh_2019]);
  spreadsheet.getRange('R21C2').setValue(array_rozvrh[0][19+posun_rozvrh_2019]);
  
  spreadsheet.getRange('R22C2').setValue(array_rozvrh[0][20+posun_rozvrh_2019]);
  spreadsheet.getRange('R23C2').setValue(array_rozvrh[0][21+posun_rozvrh_2019]);

  spreadsheet.getRange('R24C2').setValue(array_rozvrh[0][2]);  
  spreadsheet.getRange('R26C2').setValue(array_rozvrh[0][22+posun_rozvrh_2019]);//píseň 3
  spreadsheet.getRange('R27C2').setValue(array_rozvrh[0][23+posun_rozvrh_2019]);
  
  //spreadsheet.getRange('R22C3').setValue("Upratovanie skupina: " + array_rozvrh[0][posledni_sloupec_rozvrh-1]); //nastav oznámení - upratovanie - vyhozeno na základě nových pokynů S-38. Přepsalo by celou buňku)
  spreadsheet.getRange('R25C2').setValue("Čítanie: " + array_rozvrh[1][6] + ", 1: " + array_rozvrh[1][8] + ", 2: " + array_rozvrh[1][11] + ", 3: " + array_rozvrh[1][14] + ", 4: " + array_rozvrh[1][17]); //úlohy příští týden - 2019 přidána úloha 4

  //vypsání úloh z programu ŽaS pro neslyš + zaboldování všeho před závorkou (2019 přidána úloha 4)
  var zacatek_uloha1 = uloha1.split("(")[0];
  zacatek_uloha1 = zacatek_uloha1.length;
  var zacatek_uloha2 = uloha2.split("(")[0];
  zacatek_uloha2 = zacatek_uloha2.length;
  var zacatek_uloha3 = uloha3.split("(")[0];
  zacatek_uloha3 = zacatek_uloha3.length;
  var zacatek_uloha4 = uloha4.split("(")[0];
  zacatek_uloha4 = zacatek_uloha4.length;
  
  spreadsheet.getRange('R10C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText(uloha1)
  .setTextStyle(0, zacatek_uloha1, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());

  spreadsheet.getRange('R12C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText(uloha2)
  .setTextStyle(0, zacatek_uloha2, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());
   
  if (uloha3 == "-" || uloha3 == "") {
   spreadsheet.getRange('R14C1').setValue(uloha3);
  } else {
   spreadsheet.getRange('R14C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
   .setText(uloha3)
   .setTextStyle(0, zacatek_uloha3, SpreadsheetApp.newTextStyle()
   .setBold(true)
   .build())
   .build());
  }

  if (uloha4 == "-" || uloha4 == "") {
   spreadsheet.getRange('R16C1').setValue(uloha4);
  } else {  
   spreadsheet.getRange('R16C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
   .setText(uloha4)
   .setTextStyle(0, zacatek_uloha4, SpreadsheetApp.newTextStyle()
   .setBold(true)
   .build())
   .build());
  }
  
  //vypsání bodů z WOL
  spreadsheet.getRange('R6C3').setValue(obsah_prejav.split("(")[0].substr(2)); //substr ořízne 2 mezery ze začátku
  
  obsah_ctenibi = "  Čítanie Biblie " + obsah_ctenibi; //úprava na 2019
  
  var zacatek_ctenibi = obsah_ctenibi.split("(")[0];
  zacatek_ctenibi = zacatek_ctenibi.length - 2;
  
  spreadsheet.getRange('R8C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText(obsah_ctenibi.substr(2))
  .setTextStyle(0, zacatek_ctenibi, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());  
  
  var leva_program1 = obsah_program1.split(")")[0] + ")";
  
  var zacatek_leva_program1 = leva_program1.split('(')[0];
  zacatek_leva_program1 = zacatek_leva_program1.length - 2;
  
  var prava_program1 = obsah_program1.split(")")[1];

  //2019 - všechna další R zvýšena o 2 (kromě týdenního čtení)
  spreadsheet.getRange('R20C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText(leva_program1.substr(2))
  .setTextStyle(0, zacatek_leva_program1, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());
  
  spreadsheet.getRange('R20C3').setValue(prava_program1.substr(2));
  
  if (dvaprogramy == true) {
    
    var leva_program2 = obsah_program2.split(")")[0] + ")";
  
    var zacatek_leva_program2 = leva_program2.split('(')[0];
    zacatek_leva_program2 = zacatek_leva_program2.length - 2;
  
    var prava_program2 = obsah_program2.split(")")[1];

    spreadsheet.getRange('R21C1').setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText(leva_program2.substr(2))
    .setTextStyle(0, zacatek_leva_program2, SpreadsheetApp.newTextStyle()
    .setBold(true)
    .build())
    .build());
  
   spreadsheet.getRange('R21C3').setValue(prava_program2.substr(2));
      
  } else {
    
   spreadsheet.getRange('R21C1').setValue("-");
   spreadsheet.getRange('R21C2').setValue("-");
   spreadsheet.getRange('R21C3').setValue("-");
    
  } 
  
  spreadsheet.getRange('R22C3').setValue(obsah_stubi.split(")")[1].substr(2)); //substr ořízne 2 mezery ze začátku  
 
  spreadsheet.getRange('R4C3').setValue(tydenni_cteni);
  
}

//-------------------------- PŘÍPRAVA ROZVRHU NA TISK ------------------

function Priprav_rozvrh_na_tisk() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  //kontrola, jestli nestojíme kurzorem mimo sloupec A
  if (spreadsheet.getActiveCell().getColumn()>1) {
     Browser.msgBox("Nestojíš v sloupci A, končím.", Browser.Buttons.OK_CANCEL);
     return;
  }
  
  //zjištění datumu
  var datum = spreadsheet.getCurrentCell().getValue();
  var mesic = datum.getMonth()+1;
  var rok = datum.getFullYear();
  var den = datum.getDate();
  
  //zjištění počtu týdnů v měsíci
  var pocet_pondelku = Pondelky(rok, mesic-1); //0=leden, 11=prosinec, proto odečítáme 1
  
  //range s rozvrhem do array
  //nalezení posledního sloupce rozvrhu
  var posledni_sloupec_rozvrh = spreadsheet.getLastColumn();
  //nalezení současného řádku
  var soucasny_radek_rozvrh = spreadsheet.getActiveCell().getRow();
  var posledni_radek_mesic = soucasny_radek_rozvrh + pocet_pondelku - 1;
  //var array_rozvrh = spreadsheet.getRange('A33:AK36').getValues();
  var array_rozvrh = spreadsheet.getRange('R' + soucasny_radek_rozvrh + 'C1:R' + posledni_radek_mesic + 'C' + posledni_sloupec_rozvrh).getValues();
  
  //upraveni šablony rozvrhu na počet týdnů
  if (pocet_pondelku == 4) {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('(Tisk rozvrhu)'), true);
    spreadsheet.getRange('I:J').activate();
    spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
     
    //vložení ikonek na střed
    spreadsheet.getRange("D6").setHorizontalAlignment('left').setVerticalAlignment('middle');
    spreadsheet.getRange("D10").setHorizontalAlignment('left').setVerticalAlignment('middle');
    spreadsheet.getRange("D22").setHorizontalAlignment('left').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
     spreadsheet.getRange("D35").setHorizontalAlignment('left').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
    spreadsheet.getRange("D42").setHorizontalAlignment('left').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
    
    //NEMAZAT - 2 metody VLOŽENÍ OBRÁZKU
    //var getImageBlob = DriveApp.getFileById("18r5_r05u6u6HqRGrkzYH7UHYcau1U3KN").getBlob();
    //spreadsheet.getActiveSheet().insertImage(getImageBlob,4,6,24,3); //insertImage(blob, column, row, offsetX, offsetY)
    //spreadsheet.getRange("D6").setValue('=IMAGE("https://docs.google.com/uc?export=download&id=18r5_r05u6u6HqRGrkzYH7UHYcau1U3KN"; 3)');
    //spreadsheet.getRange("D6").setHorizontalAlignment('center').setVerticalAlignment('middle');  
    
  } else {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('(Tisk rozvrhu)'), true);
    spreadsheet.getActiveSheet().showColumns(9);
    spreadsheet.getActiveSheet().showColumns(10);
    
    //vložení ikonek na střed
    spreadsheet.getRange("D6").setHorizontalAlignment('right').setVerticalAlignment('middle');
    spreadsheet.getRange("D10").setHorizontalAlignment('right').setVerticalAlignment('middle');
    spreadsheet.getRange("D22").setHorizontalAlignment('right').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
     spreadsheet.getRange("D35").setHorizontalAlignment('right').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
    spreadsheet.getRange("D42").setHorizontalAlignment('right').setVerticalAlignment('middle'); //od 2019 zvětšeno o 3
    
  }
  
  //vložení titulku
  spreadsheet.getRange("A1").setValue('Program zhromaždení       ' + mesic + '/' + rok);

  //přečtení nastavení který den je shromáždění přes týden a kdy přes víkend
  var den_tyden = spreadsheet.getSheetByName("Nastavení").getRange('B3').getValues();
  var den_vikend = spreadsheet.getSheetByName("Nastavení").getRange('B5').getValues();
  
  //vypsání arraye s rozvrhem
  //array je takto: array_rozvrh[řádek od 0][sloupec od 0]
  for (var i = 0; i < pocet_pondelku; ++i) {
     
    var sloupec = (1 + i) * 2; //vynechává úzké sloupce
    
    var datum_tydenniho = array_rozvrh[i][0];
    datum_tydenniho.setDate(datum_tydenniho.getDate() + Number(den_tyden));
   
    spreadsheet.getRange('R3C' + sloupec).setValue(datum_tydenniho);
    
    spreadsheet.getRange('R4C' + sloupec).setValue(array_rozvrh[i][1]);
    spreadsheet.getRange('R5C' + sloupec).setValue(array_rozvrh[i][2]); 
    spreadsheet.getRange('R7C' + sloupec).setValue(array_rozvrh[i][4]);
    spreadsheet.getRange('R8C' + sloupec).setValue(array_rozvrh[i][5]);
    spreadsheet.getRange('R9C' + sloupec).setValue(array_rozvrh[i][6]);  
    spreadsheet.getRange('R11C' + sloupec).setValue(array_rozvrh[i][8]);
    spreadsheet.getRange('R12C' + sloupec).setValue(array_rozvrh[i][10]);
    spreadsheet.getRange('R14C' + sloupec).setValue(array_rozvrh[i][11]);
    spreadsheet.getRange('R15C' + sloupec).setValue(array_rozvrh[i][13]);
    spreadsheet.getRange('R17C' + sloupec).setValue(array_rozvrh[i][14]);
    spreadsheet.getRange('R18C' + sloupec).setValue(array_rozvrh[i][16]);
    
    spreadsheet.getRange('R20C' + sloupec).setValue(array_rozvrh[i][17]); //od 2019 úloha č. 4
    spreadsheet.getRange('R21C' + sloupec).setValue(array_rozvrh[i][19]); //od 2019 úloha č. 4   
        
    spreadsheet.getRange('R23C' + sloupec).setValue(array_rozvrh[i][20]); //od 2019 všechny R dále jsou o 3 větší, všechny čísla v [] jsou o 3 větší
    spreadsheet.getRange('R25C' + sloupec).setValue(array_rozvrh[i][21]);
    spreadsheet.getRange('R26C' + sloupec).setValue(array_rozvrh[i][22]);
    spreadsheet.getRange('R28C' + sloupec).setValue(array_rozvrh[i][23]);
    spreadsheet.getRange('R29C' + sloupec).setValue(array_rozvrh[i][24]);
    spreadsheet.getRange('R31C' + sloupec).setValue(array_rozvrh[i][25]);
    spreadsheet.getRange('R32C' + sloupec).setValue(array_rozvrh[i][26]);
    
    var datum_vikendoveho = array_rozvrh[i][0];
    datum_vikendoveho.setDate(datum_vikendoveho.getDate() + Number(den_vikend) - Number(den_tyden)); //datum uložené v array se mění, tak ho měníme zpátky a pak přidáváme
    
    spreadsheet.getRange('R34C' + sloupec).setValue(datum_vikendoveho);
  
    spreadsheet.getRange('R36C' + sloupec).setValue(array_rozvrh[i][27]);
    spreadsheet.getRange('R37C' + sloupec).setValue(array_rozvrh[i][29]);
    spreadsheet.getRange('R39C' + sloupec).setValue(array_rozvrh[i][30]);
    spreadsheet.getRange('R40C' + sloupec).setValue(array_rozvrh[i][31]);
    spreadsheet.getRange('R41C' + sloupec).setValue(array_rozvrh[i][32]);
    spreadsheet.getRange('R43C' + sloupec).setValue(array_rozvrh[i][33]);
    spreadsheet.getRange('R44C' + sloupec).setValue(array_rozvrh[i][34]);
    spreadsheet.getRange('R45C' + sloupec).setValue(array_rozvrh[i][35]);
    spreadsheet.getRange('R47C' + sloupec).setValue(array_rozvrh[i][37]);
    spreadsheet.getRange('R48C' + sloupec).setValue(array_rozvrh[i][38]);
    spreadsheet.getRange('R50C' + sloupec).setValue(array_rozvrh[i][39]);
    
  }

  
}

//-------------------------- VYGENEROVÁNÍ 1 ŘÁDKU ROZVRHU ------------------

function Pridel_radek() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  //kontrola, jestli nestojíme kurzorem mimo sloupec A
  if (spreadsheet.getActiveCell().getColumn()>1) {
     Browser.msgBox("Nestojíš v sloupci A, končím.", Browser.Buttons.OK_CANCEL);
     return;
  }
  //zjištění čísla posledního řádku v hárku "Osoby"
  var posledni_radek = spreadsheet.getSheetByName("Osoby").getDataRange().getValues().length;
  //zjištění datumu
  var datum = spreadsheet.getCurrentCell().getValue();
  var mesic = datum.getMonth()+1;
  var rok = datum.getFullYear();
  var den = datum.getDate();
  
  //který týden je to v měsíci? (počítáno od prvního pondělku) (přidáno 2019)
  var ktery_tyden = Kolikaty_tyden(rok, mesic-1,datum); //0=leden, 11=prosinec, proto odečítáme 1  
 
  
  //převodník na písně (update 2020 - sjj)
  //pozice v array je číslo písně, číslo uložené pod touto pozicí je alternativní píseň
  var prevodnik_pisne = [];
  prevodnik_pisne[0] = 0;
  prevodnik_pisne[1] = 1;
  prevodnik_pisne[2] = 1;  
  prevodnik_pisne[3] = 23;
  prevodnik_pisne[4] = 22;
  prevodnik_pisne[5] = 15;
  prevodnik_pisne[6] = 15;
  prevodnik_pisne[7] = 23;
  prevodnik_pisne[8] = 49;
  prevodnik_pisne[9] = 46;
  prevodnik_pisne[10] = 9;
  prevodnik_pisne[11] = 15;
  prevodnik_pisne[12] = 2;
  prevodnik_pisne[13] = 5;
  prevodnik_pisne[14] = 30;
  prevodnik_pisne[15] = 5;
  prevodnik_pisne[16] = 14;
  prevodnik_pisne[17] = 25;
  prevodnik_pisne[18] = 5;
  prevodnik_pisne[19] = 8;
  prevodnik_pisne[20] = 2;
  prevodnik_pisne[21] = 40;
  prevodnik_pisne[22] = 16;
  prevodnik_pisne[23] = 30;
  prevodnik_pisne[24] = 16;
  prevodnik_pisne[25] = 46;
  prevodnik_pisne[26] = 53;
  prevodnik_pisne[27] = 9;
  prevodnik_pisne[28] = 27;
  prevodnik_pisne[29] = 34;
  prevodnik_pisne[30] = 51;
  prevodnik_pisne[31] = 26;
  prevodnik_pisne[32] = 27;
  prevodnik_pisne[33] = 38;
  prevodnik_pisne[34] = 29;
  prevodnik_pisne[35] = 42;
  prevodnik_pisne[36] = 52;
  prevodnik_pisne[37] = 10;
  prevodnik_pisne[38] = 23;
  prevodnik_pisne[39] = 4;
  prevodnik_pisne[40] = 31;
  prevodnik_pisne[41] = 6;
  prevodnik_pisne[42] = 6;
  prevodnik_pisne[43] = 13;
  prevodnik_pisne[44] = 38;
  prevodnik_pisne[45] = 22;
  prevodnik_pisne[46] = 2;
  prevodnik_pisne[47] = 51;
  prevodnik_pisne[48] = 48;
  prevodnik_pisne[49] = 11;
  prevodnik_pisne[50] = 48;
  prevodnik_pisne[51] = 7;
  prevodnik_pisne[52] = 7;
  prevodnik_pisne[53] = 45;
  prevodnik_pisne[54] = 32;
  prevodnik_pisne[55] = 33;
  prevodnik_pisne[56] = 34;
  prevodnik_pisne[57] = 18;
  prevodnik_pisne[58] = 47;
  prevodnik_pisne[59] = 9;
  prevodnik_pisne[60] = 10;
  prevodnik_pisne[61] = 17;
  prevodnik_pisne[62] = 28;
  prevodnik_pisne[63] = 31;
  prevodnik_pisne[64] = 44;
  prevodnik_pisne[65] = 45;
  prevodnik_pisne[66] = 47;
  prevodnik_pisne[67] = 47;
  prevodnik_pisne[68] = 44;
  prevodnik_pisne[69] = 17;
  prevodnik_pisne[70] = 44; 
  prevodnik_pisne[71] = 17;
  prevodnik_pisne[72] = 10;
  prevodnik_pisne[73] = 33;
  prevodnik_pisne[74] = 28;
  prevodnik_pisne[75] = 10;
  prevodnik_pisne[76] = 16;
  prevodnik_pisne[77] = 47;
  prevodnik_pisne[78] = 40;
  prevodnik_pisne[79] = 7;
  prevodnik_pisne[80] = 1;  
  prevodnik_pisne[81] = 44;
  prevodnik_pisne[82] = 45;
  prevodnik_pisne[83] = 45;
  prevodnik_pisne[84] = 40;
  prevodnik_pisne[85] = 21;
  prevodnik_pisne[86] = 20;
  prevodnik_pisne[87] = 20;
  prevodnik_pisne[88] = 11;
  prevodnik_pisne[89] = 6;
  prevodnik_pisne[90] = 53;
  prevodnik_pisne[91] = 13;
  prevodnik_pisne[92] = 13;
  prevodnik_pisne[93] = 20;
  prevodnik_pisne[94] = 37;
  prevodnik_pisne[95] = 43;
  prevodnik_pisne[96] = 37;
  prevodnik_pisne[97] = 48;
  prevodnik_pisne[98] = 37;
  prevodnik_pisne[99] = 31;
  prevodnik_pisne[100] = 50;
  prevodnik_pisne[101] = 53;
  prevodnik_pisne[102] = 42;
  prevodnik_pisne[103] = 42;
  prevodnik_pisne[104] = 38;
  prevodnik_pisne[105] = 3;
  prevodnik_pisne[106] = 3;
  prevodnik_pisne[107] = 50;
  prevodnik_pisne[108] = 18;
  prevodnik_pisne[109] = 25;
  prevodnik_pisne[110] = 14;
  prevodnik_pisne[111] = 28;
  prevodnik_pisne[112] = 39;
  prevodnik_pisne[113] = 39;
  prevodnik_pisne[114] = 35;
  prevodnik_pisne[115] = 35;
  prevodnik_pisne[116] = 50;
  prevodnik_pisne[117] = 19;
  prevodnik_pisne[118] = 54;
  prevodnik_pisne[119] = 54;
  prevodnik_pisne[120] = 21;
  prevodnik_pisne[121] = 52;
  prevodnik_pisne[122] = 32;
  prevodnik_pisne[123] = 43;
  prevodnik_pisne[124] = 18;
  prevodnik_pisne[125] = 21;
  prevodnik_pisne[126] = 43;
  prevodnik_pisne[127] = 29;
  prevodnik_pisne[128] = 24;
  prevodnik_pisne[129] = 51;
  prevodnik_pisne[130] = 35;
  prevodnik_pisne[131] = 36;
  prevodnik_pisne[132] = 36;
  prevodnik_pisne[133] = 41;
  prevodnik_pisne[134] = 41;
  prevodnik_pisne[135] = 11;
  prevodnik_pisne[136] = 26;
  prevodnik_pisne[137] = 3;
  prevodnik_pisne[138] = 4;
  prevodnik_pisne[139] = 55;
  prevodnik_pisne[140] = 55;
  prevodnik_pisne[141] = 1;
  prevodnik_pisne[142] = 54;
  prevodnik_pisne[143] = 32;
  prevodnik_pisne[144] = 24;
  prevodnik_pisne[145] = 19;
  prevodnik_pisne[146] = 14;
  prevodnik_pisne[147] = 12;
  prevodnik_pisne[148] = 33;
  prevodnik_pisne[149] = 46;
  prevodnik_pisne[150] = 49;
  prevodnik_pisne[151] = 12;
  
  //-------------- vytáhnutí písní ŽaS z WOL -----------------
  //fetchnutí WOL s rozvrhem na daný den
  var url = 'https://wol.jw.org/sk/wol/dt/r38/lp-v/' + rok + '/' + mesic + '/' + den;
  
  var html = UrlFetchApp.fetch(url).getContentText(); 
  //nalezení 3 pozic písní
  var vyskyty = [];
  for (i = 0; i < html.length; ++i) {
    if (html.substring(i, i + 'Pieseň č.'.length) == 'Pieseň č.') {
      vyskyty.push(i);
    }
  }
  
  //uložení 3 stringů
  var pisnicky_zas = [];
  for (i=0; i<3; ++i){
    var vytahnuto = html.substring(vyskyty[i], vyskyty[i] + 13);
    
    //oříznutí ze začátku
    vytahnuto = vytahnuto.substr(10);
    //oříznutí na konci
    vytahnuto = vytahnuto.split("<")[0];
      
    //převedení písně na rozsah 55
    vytahnuto = prevodnik_pisne[vytahnuto];
    pisnicky_zas.push(vytahnuto);
  }
  //vysledek: pisnicky_zas (array se třemi písněmi)
  
  //-------------- písně SV -----------------
  //nalezení URL SV
  //vyextrahování všech inner hyperlinků z html
  var inner_links_arr= [];
  var linkRegExp = /href="(.*?)"/gi; // regex expression object
  var match = linkRegExp.exec(html);
  while (match != null) { // we filter only inner links and not pdf docs
    if (match[1].indexOf('#') !== 0
      && match[1].indexOf('http://') !== 0
      && match[1].indexOf('https://') !== 0
      && match[1].indexOf('mailto:') !== 0
      && match[1].indexOf('.pdf') === -1 ) 
      {
        inner_links_arr.push(match[1]);
      }
   match = linkRegExp.exec(html);
  }
  //link na SV je čtvrtý odzadu 
  var link_na_sv = "https://wol.jw.org" + inner_links_arr[inner_links_arr.length-4];
  
  //fetchnutí WOL s písněmi na SV
  var html_sv = UrlFetchApp.fetch(link_na_sv).getContentText(); 
  
  //nalezení piesne 1
  for (i = 0; i < html_sv.length; ++i) {
    if (html_sv.substring(i, i + 'PIESEŇ'.length) == 'PIESEŇ') {
      var vyskyty_pi1 = i;
      break; //vyskočí po prvním nálezu
    }
  }
    
  //uložení stringu s písní 1
  var pisnicka_sv1 = html_sv.substring(vyskyty_pi1, vyskyty_pi1 + 13);
  //oříznutí ze začátku
  pisnicka_sv1 = pisnicka_sv1.substr(10);
  //oříznutí na konci
  pisnicka_sv1 = pisnicka_sv1.split("<")[0]
  //převedení písně na rozsah 55
  pisnicka_sv1 = prevodnik_pisne[pisnicka_sv1];
  //vysledek: pisnicka_sv1
  
  //nalezení piesne 2
    for (i = 0; i < html_sv.length; ++i) {
    if (html_sv.substring(i, i + 'PIESEŇ'.length) == 'PIESEŇ') {
      var vyskyty_pi2 = i; //druhý nález přepíše první
    }
  }
  
  //uložení stringu s písní 2
  var pisnicka_sv2 = html_sv.substring(vyskyty_pi2, vyskyty_pi2 + 13);
  //oříznutí ze začátku
  pisnicka_sv2 = pisnicka_sv2.substr(10);
  //oříznutí na konci
  pisnicka_sv2 = pisnicka_sv2.split("<")[0]
  //oříznutí na začátku - zrušeno 2019
  //if (pisnicka_sv2.indexOf(">") !== -1) {
  // pisnicka_sv2 = pisnicka_sv2.split(">")[1];
  //}
  
  //převedení písně na rozsah 55
  pisnicka_sv2 = prevodnik_pisne[pisnicka_sv2];
  
  //přečtení nastavení který den je shromáždění přes týden a kdy přes víkend
  var den_tyden = spreadsheet.getSheetByName("Nastavení").getRange('B3').getValues();
  var den_vikend = spreadsheet.getSheetByName("Nastavení").getRange('B5').getValues();
  //přidání 1 dne k datumu (zjištění datumu týdenního shromáždění) - pokud je v úterý, tak 1...a v sobotu, tak 5
  var datum_tydenniho = spreadsheet.getCurrentCell().getValue();
  datum_tydenniho.setDate(datum_tydenniho.getDate() + Number(den_tyden));
  var datum_vikendoveho = spreadsheet.getCurrentCell().getValue();
  datum_vikendoveho.setDate(datum_vikendoveho.getDate() + Number(den_vikend));
  //otevření kalendáře starších
  var spreadkalendar = SpreadsheetApp.openById("[calendar file id]");
  //zpracování kalendáře
  var posledni_radek_kalendar = spreadkalendar.getSheetByName("Kalendár ST").getDataRange().getValues().length;
  var sloupec_datumy_kalendar = spreadkalendar.getRange('A2:A' + posledni_radek_kalendar).getValues();
  //číslo řádku v kalendáři, na kterém je stejné datum
  for(var n in sloupec_datumy_kalendar){
   if(sloupec_datumy_kalendar[n][0].getTime() == datum_tydenniho.getTime()) {
    var radek_v_kalendari = 2 + Number(n);
    break;
   }
  }
  //číslo řádku v kalendáři, na kterém je stejné datum - víkend
  for(var n in sloupec_datumy_kalendar){
   if(sloupec_datumy_kalendar[n][0].getTime() == datum_vikendoveho.getTime()) {
    var radek_v_kalendari_vikend = 2 + Number(n);
    break;
   }
  }
  
  //počet dovolenkujících v kalendáři
  //číslo posledního sloupce v kalendáři
  var posledni_sloupec_kalendar = spreadkalendar.getSheetByName("Kalendár ST").getDataRange().getLastColumn();
  //první řádek kalendáře do array (první číslo je číslo řádku, poslední číslo je počet sloupců)
  var prvni_radek_kalendare = spreadkalendar.getSheetByName("Kalendár ST").getRange(1,1,1,posledni_sloupec_kalendar).getValues();
  for(var n in prvni_radek_kalendare[0]){
   if(prvni_radek_kalendare[0][n] == "Poznámky") {
    var pocet_dovolenkujucich = Number(n) - 1;
    break;
   }
  }
  
  //vytvoření arraye s dovolenkujúcimi, kteří tu nebudou při shromáždění přes týden
  var dovolenkujuci = new Array(pocet_dovolenkujucich-1);
  for (var i = 2; i < pocet_dovolenkujucich+2; i++){
   //čísla: row, column
   var barva = spreadkalendar.getSheetByName("Kalendár ST").getRange(radek_v_kalendari, i).getBackgrounds();
   if (barva == "#ff0000"){
    var jmeno = spreadkalendar.getSheetByName("Kalendár ST").getRange(1, i).getValues();
    jmeno = jmeno.toString(); //převod z array na string
    dovolenkujuci.push(jmeno);
   }
  }
  dovolenkujuci = dovolenkujuci.filter(Boolean);
  var kolik_dovolenkujucich_je_pryc = dovolenkujuci.length;
  
  //vytvoření arraye s dovolenkujúcimi, kteří tu nebudou při shromáždění přes víkend
  var dovolenkujuci_vikend = new Array(pocet_dovolenkujucich-1);
  for (var i = 2; i < pocet_dovolenkujucich+2; i++){
   //čísla: row, column
   var barva = spreadkalendar.getSheetByName("Kalendár ST").getRange(radek_v_kalendari_vikend, i).getBackgrounds();
   if (barva == "#ff0000"){
    var jmeno = spreadkalendar.getSheetByName("Kalendár ST").getRange(1, i).getValues();
    jmeno = jmeno.toString(); //převod z array na string
    dovolenkujuci_vikend.push(jmeno);
   }
  }
  dovolenkujuci_vikend = dovolenkujuci_vikend.filter(Boolean);
  var kolik_dovolenkujucich_je_pryc_vikend = dovolenkujuci_vikend.length;
  
  //otevření přednášek
  var spreadprednasky = SpreadsheetApp.openById("[talks sheet id]");
  //zpracování přednášek
  var posledni_radek_prednasky = spreadprednasky.getSheetByName("rozvrh").getDataRange().getValues().length;
  var sloupec_datumy_prednasky = spreadprednasky.getRange('A2:A' + posledni_radek_prednasky).getValues();
  //číslo řádku v přednáškách, na kterém je stejné datum
  for(var n in sloupec_datumy_prednasky){
   if(sloupec_datumy_prednasky[n][0].getTime() == datum_vikendoveho.getTime()) {
    var radek_v_prednaskach = 2 + Number(n);
    break;
   }
  }   
  //přečtení jména přednášejícího a tématu přednášky a písně a jestli je tlumočník (upraveno 2019 dle toho jak to upravil Maťo)
  var prednasejici = spreadprednasky.getSheetByName("rozvrh").getRange(radek_v_prednaskach, 4).getValues();
  prednasejici = prednasejici.toString(); //převod z array na string
  var tema_prednasky = spreadprednasky.getSheetByName("rozvrh").getRange(radek_v_prednaskach, 3).getValues();
  tema_prednasky = tema_prednasky.toString(); //převod z array na string
  var prednaska_pisen = spreadprednasky.getSheetByName("rozvrh").getRange(radek_v_prednaskach, 6).getValues();
  prednaska_pisen = prednaska_pisen.toString(); //převod z array na string
  var tlmocnik_slovo = spreadprednasky.getSheetByName("rozvrh").getRange(radek_v_prednaskach, 5).getValues();
  tlmocnik_slovo = tlmocnik_slovo.toString(); //převod z array na string
  //2019 z "" upraveno na "false"
  if (tlmocnik_slovo == "false"){
    var tlmocnik = false;
  } else {
    var tlmocnik = true;
  } 
   
  //zparsování obsahů programu jedna a dva

  //odříznutí celé patičky
  var ocistene_html = html;
  var pozice_paticky = ocistene_html.indexOf('<div id="regionFooter');
  ocistene_html = ocistene_html.slice(0, pozice_paticky);
  //odříznutí celého úvodu
  var pozice_uvodu = ocistene_html.indexOf('<div id="regionMain');
  ocistene_html = ocistene_html.slice(pozice_uvodu);  
  
  //nalezení pozic bodů ul s class "noMarker"
  var vyskyty_nomarker = [];
  for (i = 0; i < ocistene_html.length; ++i) {
    if (ocistene_html.substring(i, i + 'noMarker'.length) == 'noMarker') {
      vyskyty_nomarker.push(i-11);
    }
  }  
  //počet ul s class noMarker
  var pocet_nomarker = vyskyty_nomarker.length;
  
  //1) vyhodíme všechny tyto uly s class noMarker
  //nalezení všech špatných ulů
  var spatny_ul = [];
  for (i = 0; i < pocet_nomarker; ++i) {
    //odříznutí začátku stringu
    var zbytek_ul = ocistene_html.slice(vyskyty_nomarker[i]);
    //nalezení nejbližšího uzavíracího tagu ul
    var kde_je_uzaviraci_ul = zbytek_ul.indexOf("</ul>");
    //vystřihnutí obsahu ulu
    spatny_ul.push(zbytek_ul.slice(0, kde_je_uzaviraci_ul+5));
  } 
  //vyříznutí všech špatných ulů
  for (i = 0; i < pocet_nomarker; ++i) {
    ocistene_html = ocistene_html.replace(spatny_ul[i],"");
  } 
  
  //2) hledáme <li>, které už představují jen opuntíkované body
  var vyskyty_li = [];
  for (i = 0; i < ocistene_html.length; ++i) {
    if (ocistene_html.substring(i, i + '<li>'.length) == '<li>') {
      vyskyty_li.push(i);
    }
  } 
  
  //3) pokud je vyskyty_li.length=13, je jen jeden program, pokud 14, jsou dva programy - pozn. 2019: to platí třetí a další týden. Neplatí to 1. a 2. týden. První týden platí: 12/13 a druhý týden platí 14/15.
  //-------------- program 1 nebo i 2? -----------------
  //rok 2019 - oprava podle toho, který je to týden
  var oprava2019 = 0;
  
  if (ktery_tyden == 1) {
    oprava2019 = -1;
  }
  
  if (ktery_tyden == 2) {
    oprava2019 = 1;
  } 
  
  if (vyskyty_li.length == (14+oprava2019)){
    var dvaprogramy = true;
  } else {
    var dvaprogramy = false;
  }
  //vysledek: dvaprogramy (pokud je true, tak jsou dva programy)
  
  //získáme obsah programů
  var obsah_program1 = ocistene_html.slice(vyskyty_li[9+oprava2019]);
  var kde_je_uzaviraci = obsah_program1.indexOf("</li>");
  obsah_program1 = obsah_program1.slice(0, kde_je_uzaviraci+5);   
  //převod HTML na plaintext
  obsah_program1=obsah_program1.replace(/<[^>]*>/g, '');
  
  if (dvaprogramy == true) {
    var obsah_program2 = ocistene_html.slice(vyskyty_li[10+oprava2019]);
    var kde_je_uzaviraci = obsah_program2.indexOf("</li>");
    obsah_program2 = obsah_program2.slice(0, kde_je_uzaviraci+5);   
    //převod HTML na plaintext
    obsah_program2=obsah_program2.replace(/<[^>]*>/g, '');
  }

  // ------ přidání komentáře s dovolenkami k datumu ------
  if (dovolenkujuci.length > 0 && dovolenkujuci_vikend.length > 0) {
    spreadsheet.getCurrentCell().setNote("Týden pryč: " + dovolenkujuci + "\n\n" + "Víkend pryč: " + dovolenkujuci_vikend);
  } else if (dovolenkujuci.length > 0) {
    spreadsheet.getCurrentCell().setNote("Týden pryč: " + dovolenkujuci);
  } else if (dovolenkujuci_vikend.length > 0) {
    spreadsheet.getCurrentCell().setNote("Víkend pryč: " + dovolenkujuci_vikend);
  }
    
  // ------ vyplnění úvodní písně ------
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(pisnicky_zas[0]);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ------ vyplnění předsedajícího ŽaS ------ OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_predseda_zas_rank = spreadsheet.getRange('Osoby!Q2:Q' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  var z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_predseda_zas_rank);
  //výběr jména
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_predseda_zas_rank){
    if(sloupec_predseda_zas_rank[n][0] == i) {
           
       var n = 2 + Number(n);
       var predsedajuci_zas = spreadsheet.getRange("Osoby!B"+n).getValue(); 

        if (dovolenkujuci.indexOf(predsedajuci_zas)>-1){
          //pokud je pryč, tak ho přeskoč
          } else {
          i = nejvyssi_cislo+1;  
          break;
        
       }
     }    
    }
  }  
  //zápis jména
  spreadsheet.getCurrentCell().setValue(predsedajuci_zas);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // vyplnění modlitby (stejná osoba jako předsedající ŽaS 
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(predsedajuci_zas);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ----- vyplnění prejavu ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky prejavu, dat do arraye
  var sloupec_prejav_rank = spreadsheet.getRange('Osoby!U2:U' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_prejav_rank);
  //notace array je následující: sloupec_prejav_rank[9][0] neboli [row][column]
  //i je to, co hledáme v arrayi
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_prejav_rank){
    if(sloupec_prejav_rank[n][0] == i) {
        n = 2 + Number(n);
        var prejav = spreadsheet.getRange("Osoby!B"+n).getValue();
      
        if (i == nejvyssi_cislo) {prejav = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku  
      
        if (dovolenkujuci.indexOf(prejav)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {
      
          if (prejav == predsedajuci_zas) {
             if (z == 1) {var prvni = prejav};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;                       
          } else {
          i = nejvyssi_cislo+1;  
          break;
        }
       }
     }    
    }
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(prejav);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ------ vyplnění duch. pokladů ------ OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_poklady_rank = spreadsheet.getRange('Osoby!W2:W' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  nejvyssi_cislo = Math.max.apply(null, sloupec_poklady_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_poklady_rank){   
      if(sloupec_poklady_rank[n][0] == i) {
        var n = 2 + Number(n);
        var poklad = spreadsheet.getRange("Osoby!B"+n).getValue();        
      
        if (i == nejvyssi_cislo) {poklad = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci.indexOf(poklad)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {
        
          if (poklad == predsedajuci_zas || poklad == prejav) {
             if (z == 1) {prvni = poklad};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
          } else {
           i = nejvyssi_cislo+1; //skončit vnější smyčku s i
           break;
          }       
        }                       
      }   
    }
  }  
  //zápis jména
  spreadsheet.getCurrentCell().setValue(poklad);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ------ vyplnění střední písně ------
  spreadsheet.getCurrentCell().offset(0, 15).activate(); //v roce 2019 změněno ze 12 na 15
  spreadsheet.getCurrentCell().setValue(pisnicky_zas[1]);
  
  //při duplicitě začerveň
  if (pisnicky_zas[1] == pisnicky_zas[0]){
     spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#ff0000');
   } else { 
     spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
   }
     
  // ----- vyplnění program1 ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setNote(obsah_program1);
  //projet sloupec s ranky, dat do arraye
  var sloupec_program_rank = spreadsheet.getRange('Osoby!Y2:Y' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  nejvyssi_cislo = Math.max.apply(null, sloupec_program_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_program_rank){
    if(sloupec_program_rank[n][0] == i) {
        var n = 2 + Number(n);
        var program = spreadsheet.getRange("Osoby!B"+n).getValue();
 
        if (i == nejvyssi_cislo) {program = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci.indexOf(program)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {
      
        if (program == predsedajuci_zas || program == prejav || program == poklad) {
             if (z == 1) {prvni = program};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }  
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(program);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
   
  // ----- vyplnění program2 ----- OK (pouze pokud jsou 2 programy)
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  if (dvaprogramy){
    spreadsheet.getCurrentCell().setNote(obsah_program2);
    //projet sloupec s ranky, dat do arraye
    var sloupec_program2_rank = spreadsheet.getRange('Osoby!Y2:Y' + posledni_radek).getValues();
    //zjištění nejvyššího čísla v array
    z = 1;
    var nejvyssi_cislo = Math.max.apply(null, sloupec_program2_rank);
    for (i = 1; i < nejvyssi_cislo+1; i++){
      for(var n in sloupec_program2_rank){
      if(sloupec_program2_rank[n][0] == i) {
          var n = 2 + Number(n);
          var program2 = spreadsheet.getRange("Osoby!B"+n).getValue();
      
          if (i == nejvyssi_cislo) {program2 = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
          if (dovolenkujuci.indexOf(program2)>-1){
            //pokud je starší pryč, tak ho přeskoč
            } else {     
      
          if (program2 == predsedajuci_zas || program2 == prejav || program2 == poklad || program2 == program) {
               if (z == 1) {prvni = program2};//zapamatuj si první jméno, které nemá dovolenku
               z = z + 1;
          } else {
          i = nejvyssi_cislo+1; 
          break;
          }
        }
      }
    }
    }
    //zápis jména
    spreadsheet.getCurrentCell().setValue(program2);  
    spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  } else {
    //pokud je jen jeden program, tak aby neplatily další podmínky:
    program2 = program;
  }  

   
  // ----- vyplnění Štúdium Biblie ------- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_stubi_rank = spreadsheet.getRange('Osoby!AA2:AA' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_stubi_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_stubi_rank){
    if(sloupec_stubi_rank[n][0] == i) {
        var n = 2 + Number(n);
        var stubi = spreadsheet.getRange("Osoby!B"+n).getValue();
       
        if (i == nejvyssi_cislo) {stubi = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci.indexOf(stubi)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {          
      
        if (stubi == predsedajuci_zas || stubi == prejav || stubi == poklad || stubi == program || stubi == program2) {
             if (z == 1) {prvni = stubi};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(stubi);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ----- vyplnění čítania ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_cteni_rank = spreadsheet.getRange('Osoby!AC2:AC' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_cteni_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_cteni_rank){
    if(sloupec_cteni_rank[n][0] == i) {
        var n = 2 + Number(n);
        var cteni = spreadsheet.getRange("Osoby!B"+n).getValue();
      
        if (i == nejvyssi_cislo) {cteni = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci.indexOf(cteni)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {        
      
        if (cteni == predsedajuci_zas || cteni == prejav || cteni == poklad || cteni == program || cteni == program2 || cteni == stubi) {
             if (z == 1) {prvni = cteni};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }  
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(cteni);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
    
  // ------ vyplnění závěrečné písně ------
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(pisnicky_zas[2]);
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  //při duplicitě začerveň
  if (pisnicky_zas[2] == pisnicky_zas[1] || pisnicky_zas[2] == pisnicky_zas[0]){
     spreadsheet.getActiveRangeList().setFontColor('#ff0000');
   }
  
  // ----- vyplnění záv. modlitby ŽaS ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_modli_rank = spreadsheet.getRange('Osoby!AC2:AC' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_modli_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_modli_rank){
    if(sloupec_modli_rank[n][0] == i) {
        var n = 2 + Number(n);
        var modli = spreadsheet.getRange("Osoby!B"+n).getValue();
      
           if (i == nejvyssi_cislo) {modli = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci.indexOf(modli)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {    
            
        if (modli == predsedajuci_zas || modli == prejav || modli == poklad || modli == program || modli == program2 || modli == stubi || modli == cteni) {
             if (z == 1) {prvni = modli};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }  
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(modli);   
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center'); 
  
  //----- vložení píseň k přednášce ----- 
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  if (prednaska_pisen == ""){
    } else {
      spreadsheet.getCurrentCell().setValue(prednaska_pisen); 
      spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
    } 
  
  // ----- výběr úvodní modlitba a předsedání o víkendu ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_pred_rank = spreadsheet.getRange('Osoby!AE2:AE' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_pred_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_pred_rank){
    if(sloupec_pred_rank[n][0] == i) {
        var n = 2 + Number(n);
        var pred = spreadsheet.getRange("Osoby!B"+n).getValue();
      
            if (i == nejvyssi_cislo) {pred = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci_vikend.indexOf(pred)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {      
      
        if (pred == predsedajuci_zas || pred == prejav || pred == poklad || pred == program || pred == program2 || pred == stubi || pred == cteni|| pred == modli) {
             if (z == 1) {prvni = pred};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }  
  }
  //zápis jména
  spreadsheet.getCurrentCell().setValue(pred);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(pred);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  //----- vložení řečníka přednáška ----- 
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(prednasejici); 
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
   //----- vložení téma přednáška ----- 
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(tema_prednasky); 
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('left');
  
  //----- výběr vedenie SV ----- OK
  //projet sloupec s ranky, dat do arraye
  var sloupec_sv_rank = spreadsheet.getRange('Osoby!AI2:AI' + posledni_radek).getValues();
  for(var n in sloupec_sv_rank){
    if(sloupec_sv_rank[n][0] == 1) {
        var n = 2 + Number(n);
        var sv = spreadsheet.getRange("Osoby!B"+n).getValue();
      
        if (dovolenkujuci_vikend.indexOf(sv)>-1){
          sv = "d";
          //pokud je starší pryč, dej tam "d" jako dovolenka
        }
      
        break;    
     }
  } 
  //zatím nezapisujeme
  
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  if (tlmocnik) {
  //projet sloupec s ranky, dat do arraye
  var sloupec_tlumocnik_rank = spreadsheet.getRange('Osoby!AG2:AG' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  z = 1;
  var nejvyssi_cislo = Math.max.apply(null, sloupec_tlumocnik_rank);
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_tlumocnik_rank){
    if(sloupec_tlumocnik_rank[n][0] == i) {
        var n = 2 + Number(n);
        var tlumocnik = spreadsheet.getRange("Osoby!B"+n).getValue();
      
            if (i == nejvyssi_cislo) {tlumocnik = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci_vikend.indexOf(tlumocnik)>-1){
          //pokud je starší pryč, tak ho přeskoč
          } else {      
      
        if (tlumocnik == prednasejici || tlumocnik == sv) {
             if (z == 1) {prvni = tlumocnik};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
   }  
  }
  spreadsheet.getCurrentCell().setValue(tlumocnik);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  }
  
  // ------ vyplnění první písně SV ------
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(pisnicka_sv1);
  //při duplicitě začerveň
  if (pisnicka_sv1 == prednaska_pisen){
     spreadsheet.getActiveRangeList().setFontWeight('bold').setFontColor('#ff0000').setHorizontalAlignment('center');
   } else { 
     spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
   }
 
  //zápis jména vedenie SV
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(sv); 
  if (sv == "d"){
     spreadsheet.getActiveRangeList().setFontWeight('bold').setFontColor('#9900ff').setHorizontalAlignment('center');
   } else { 
     spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
   }
  
  // ------ vyplnění druhé písně SV ------
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(pisnicka_sv2);
  //při duplicitě začerveň
  if (pisnicka_sv2 == prednaska_pisen || pisnicka_sv2 == pisnicka_sv1){
     spreadsheet.getActiveRangeList().setFontWeight('bold').setFontColor('#ff0000').setHorizontalAlignment('center');
   } else { 
     spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
   }
  
  //----- vložení modlitby závěrečné ----- 
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(prednasejici); 
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');  
 
  // ------ výběr usporad1 a 2 ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_usp_rank = spreadsheet.getRange('Osoby!AK2:AK' + posledni_radek).getValues();
  //zjištění nejvyššího čísla v array
  var nejvyssi_cislo = Math.max.apply(null, sloupec_usp_rank);
  //výběr jména
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_usp_rank){
    if(sloupec_usp_rank[n][0] == i) {
           
       var n = 2 + Number(n);
       var usp1 = spreadsheet.getRange("Osoby!B"+n).getValue(); 

        if (dovolenkujuci_vikend.indexOf(usp1)>-1 || dovolenkujuci.indexOf(usp1)>-1){
          //pokud je pryč, tak ho přeskoč
          } else {
          i = nejvyssi_cislo+1;  
          break;
        
       }
     }    
    }
  }  
  //usporiadatel 2 
  z = 1;
  for (i = 1; i < nejvyssi_cislo+1; i++){
    for(var n in sloupec_usp_rank){
    if(sloupec_usp_rank[n][0] == i) {
        var n = 2 + Number(n);
        var usp2 = spreadsheet.getRange("Osoby!B"+n).getValue();
      
            if (i == nejvyssi_cislo) {usp2 = prvni; break;} //když už všichni byli, tak dej prvního, který nemá dovolenku        
             
        if (dovolenkujuci_vikend.indexOf(usp2)>-1 || dovolenkujuci.indexOf(usp2)>-1){
          //pokud je pryč, tak ho přeskoč
          } else {      
      
        if (usp2 == usp1) {
             if (z == 1) {prvni = usp2};//zapamatuj si první jméno, které nemá dovolenku
             z = z + 1;
        } else {
        i = nejvyssi_cislo+1; 
        break;
        }
      }
    }
  }  
  }
  
  //zápis jména
  spreadsheet.getCurrentCell().setValue(usp1); 
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(usp2); 
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  
  // ----- výběr upratovanie ----- OK
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //projet sloupec s ranky, dat do arraye
  var sloupec_uprat_rank = spreadsheet.getRange('Osoby!AM2:AM' + posledni_radek).getValues();
  for(var n in sloupec_uprat_rank){
    if(sloupec_uprat_rank[n][0] == 1) {
        n = 2 + Number(n);
        var uprat = spreadsheet.getRange("Osoby!B"+n).getValue();
        break;
    } 
  }     
  spreadsheet.getCurrentCell().setValue(uprat);  
  spreadsheet.getActiveRangeList().setFontWeight('bold').setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().offset(0, -36).activate();
  
};


