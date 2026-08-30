const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        BorderStyle, WidthType, ShadingType, HeadingLevel, PageBreak, PageOrientation } = require('docx');
const PDFDocument = require('pdfkit');

// Species data organized alphabetically with their specific gravity settings and scientific names
const speciesData = [
  { name: "African Blackwood", setting: "1.27", scientific: "Dalbergia melanoxylon", synonyms: ["African Grenadillo","Grenadilla","Mpingo"], woodDb: "african-blackwood" },
  { name: "Afrormosia", setting: "0.65", scientific: "Pericopsis elata", woodDb: "afrormosia" },
  { name: "Alder, Red", setting: "0.41", scientific: "Alnus rubra", synonyms: ["Alder","Oregon Alder","Pacific Coast Alder","Western Alder","Western Red Alder"], woodDb: "red-alder" },
  { name: "Amazakoue (Ovangkol)", setting: "0.82", scientific: "Guibourtia ehie", synonyms: ["Amazoue","Anokye","Ehie","Gabon ovanko","Ghana anokye","Hyedua","Hyedua-nini","Hyeduanini","Ivory Coast Amazakoue","Mongoy","Mongoy Walnut","Ovengkol","Pallisandro"], woodDb: "ovangkol" },
  { name: "American Red Oak", setting: "0.63", scientific: "Quercus rubra", synonyms: ["Black oak","Canadian red oak","Gray Oak","Northern Red Oak","Red Oak","Southern Red Oak"], woodDb: "red-oak" },
  { name: "Andiroba", setting: "0.57", scientific: "Carapa guianensis", woodDb: "andiroba" },
  { name: "Ash, Black", setting: "0.49", scientific: "Fraxinus nigra", synonyms: ["American Black Ash","Basket ash","Brown ash","Fresno","Hoop Ash","Swamp Ash","Water Ash"], woodDb: "black-ash" },
  { name: "Ash, Blue", setting: "0.58", scientific: "Fraxinus quadrangulata", woodDb: "blue-ash" },
  { name: "Ash, Green", setting: "0.56", scientific: "Fraxinus pennsylvanica", synonyms: ["American Ash","Darlington Ash","Fresno","Pumpkin Ash","Swamp Ash","Water Ash","Water Ash Profunda"], woodDb: "green-ash" },
  { name: "Ash, Oregon", setting: "0.55", scientific: "Fraxinus latifolia", woodDb: "oregon-ash" },
  { name: "Ash, Red", setting: "0.55", scientific: "Fraxinus pennsylvanica", synonyms: ["American Ash","Darlington Ash","Fresno","Pumpkin Ash","Swamp Ash","Water Ash","Water Ash Profunda"], woodDb: "green-ash" },
  { name: "Ash, White", setting: "0.60", scientific: "Fraxinus americana", synonyms: ["American Ash","Biltmore ash","Biltmore white ash","Canadian ash","Cane ash","Fresno","Quebec Ash","White River Ash"], woodDb: "white-ash" },
  { name: "Aspen, Bigtooth", setting: "0.39", scientific: "Populus grandidentata", woodDb: "bigtooth-aspen" },
  { name: "Aspen, Quaking", setting: "0.38", scientific: "Populus tremuloides", synonyms: ["Alamo Tremblon","Aspe","Aspen","Canadian aspen","Espe","Peuplier Tremble","Trembling Aspen","Zitterpappel"], woodDb: "quaking-aspen" },
  { name: "Avodire", setting: "0.51", scientific: "Turraeanthus africanus", synonyms: ["Adoma","African Furniture Wood","African Satinwood","Agbe","Agbouain","Agboue","Agboui","Aneadwen","Aniadwen","Anyadwe","Apapaye","Apata","Apaya","Apeya","Appayia","Assama","Avodere","Blimah-pu","Blimahpu","Duabedom","Engan","Enyaadwene","Gakrada","Guaso","Gyakrada","Hadwe","Hague","Hakue","Kakne","Kisanda","Kwadwema","Lusambar","M'fube","Oduma","Olon","Songo","Sunkroasu","Wansenwa","White Mahohany","Wogya","Wonzanwa"], woodDb: "avodire" },
  { name: "Baldcypress", setting: "0.46", scientific: "Taxodium distichum", synonyms: ["Ahuehuete","Ahuehuetl","Bald cypress","Black cypress","Buck cypress","Cipres de Montezuma","Cipreso","Common bald cypress","Cow cypress","Cypress","Faux Satine","Florida Cypress","Gulf Cypress","Inland Cypress","Louisiana Cypress","Louisiana Red Cypress","Pentamon","Pentamu","Pond Cypress","Red Cypress","Sabino","Southern Cypress","Swamp Cypress","Tidewater Red Cypress","Tnuyuca","White Cypress","Xina","Yaga- Chichilino","Yaga-Guichi","Yellow Cypress","Yuca-Ndatura"], woodDb: "bald-cypress" },
  { name: "Balsa", setting: "0.14", scientific: "Ochroma pyramidale", woodDb: "balsa" },
  { name: "Balsamo (Myroxylon)", setting: "0.95", scientific: "Myroxylon balsamum", synonyms: ["Cabreuva", "Estoraque", "Incienso", "Quina"], woodDb: "santos-mahogany" },
  { name: "Balsamo (Protium)", setting: "0.55", scientific: "Protium spp." },
  { name: "Banak (Virola)", setting: "0.45", scientific: "Virola spp." },
  { name: "Basswood, American", setting: "0.37", scientific: "Tilia americana", synonyms: ["American Lime","Basswood","Bee tree","Beetree","Beetree linden","Carolina linden","Florida Basswood","Florida Linden","Limetree","Linden","Linn","White Basswood","Whitewood","Willow"], woodDb: "basswood" },
  { name: "Beech, American", setting: "0.64", scientific: "Fagus grandifolia", synonyms: ["Beech","Canadian beech","Red Beech"], woodDb: "american-beech" },
  { name: "Beech, Euro", setting: "0.67", scientific: "Fagus sylvatica", synonyms: ["Beech","Carpathian beech","Common beech","Danish beech","English beech","European beech","Fayard","French beech","Japanese Beech","Rod Bok","Romanian Beech","Rumanian Beech","Slavonian Beech","Varlig Bok","Yugoslavian beech"], woodDb: "european-beech" },
  { name: "Benge", setting: "0.70", scientific: "Guibourtia arnoldiana" },
  { name: "Birch, Paper", setting: "0.55", scientific: "Betula papyrifera", synonyms: ["American Birch","American White Birch","Birch","Black birch","Canadian white birch","Canoe birch","Kenai Birch","Kenai Paper Birch","Paper Bbirch","Paper Canoe","Red Birch","Silver Birch","Western Paper Birch","Western White Birch"], woodDb: "paper-birch" },
  { name: "Birch, Sweet", setting: "0.65", scientific: "Betula lenta", synonyms: ["Birch","Black birch","Cherry birch"], woodDb: "sweet-birch" },
  { name: "Birch, White", setting: "0.53", scientific: "Betula populifolia", synonyms: ["Gray Birch","Grey Birch"], woodDb: "gray-birch" },
  { name: "Birch, Yellow", setting: "0.62", scientific: "Betula alleghaniensis", synonyms: ["American Birch","Betula wood","Birch","Black birch","Canadian silky wood","Canadian yellow birch","Curly Birch","Gold Birch","Gray Birch","Hard Birch","Quebec Birch","Silver Birch","Swamp Birch"], woodDb: "yellow-birch" },
  { name: "Bolivian Rosewood", setting: "0.78", scientific: "Machaerium scleroxylon", synonyms: ["Caviuna", "Jacaranda Pardo", "Morado", "Pau Ferro", "Santos Rosewood"], woodDb: "pau-ferro" },
  { name: "Box", setting: "0.83", scientific: "Buxus sempervirens", synonyms: ["Abassian Boxwood","Assame-tsuge","Bois commun","Boxwood","Chikri","Circassian boxwood","Common box","European boxwood","Indian Boxwood","Iranian Boxwood","Japanese Box","North African Boxwood","Papar","Paprang","Papri","Persian Boxwood","Shamshad","Shanda Laghune","Shibsashin","Shumaj","True Box","Turkey Boxwood","Turkish Boxwood"], woodDb: "boxwood" },
  { name: "Brazilian Cherry", setting: "0.83", scientific: "Hymenaea courbaril", synonyms: ["Abati","Abati Copal do Brasil","Abati Timbary","Alga","Algarobo","Algarrobo","Animebaum","Arati","Arvore Copal","Asucar-juain","Avati","Azucar huayo","Bati timbary","Bois de courabaril","Bois de courabil","Bois de courbaril","Bois de simire","Brazilian gum-copal tree","Cacachien","Cachien","Caguairan","Cannariboom","Caouroubali","Caroubier","Caroubier de la Guyane","Cataqui-iamani","Chimidida","Cimiri","Ciruelo","Coapinol","Coapinoloe","Colorado","Comer de arara","Copal","Copalier","Copalier d'Amerique","Copalier de Amerique","Copinol","Copinole","Corobore","Coubaril de savane","Courabil","Courabil hout","Courbaril","Courbaril de savane","Courbaril montagne","Courbaril plum","Courbarilhout","Cuapinol","Cuapinole","Cuapinoli","Cupainol","Curbaril","Diphylle pois confiture","Diphylle pois de confiture","Gaupinol","Gom anime boom","Gomme armimec","Guapinol","Guapinole","Gum-anime tree","Henschreckenbaum","Henschrenkenbaum","Iibiuva","Itaiba","Jassai","Jataby","Jatahy","Jatahy Cafe","Jatahy Peba","Jatahy Roxa","Jatahy Roxo","Jatai","Jatai Monde","Jatai Uba","Jatai-Acu","Jataiba","Jataizinho","Jatany","Jatauba","Jatay","Jatei","Jatel","Jatoba","Jatoba de Anta","Jatoba de Porco","Jatoba Roxa","Jatoba Trapuca","Jatoba Verdadeiro","Jatobs","Jatuba","Jengi Kanda","Jetahy","Jetahy Accu","Jetahy Preta","Jetahy Roxo","Jetai","Jetai de Pernambuco","Jetaiba","Jetaici","Jetay","Jetoba Roxo","Jetui Peba","Julchihout","Jupati","Jutaahy Assu","Jutahy","Jutahy Assu","Jutahy Cafe","Jutahy Catina","Jutahy Catinga","Jutahy da Varzea","Jutahy do Campo","Jutahy do Igapo","Jutahy Miry","Jutahy Peba","Jutahy Pororoca","Jutahy Roxo","Jutai","Jutai Branco","Jutai Cafe","Jutai Catinga","Jutai da Varzea","Jutai do Campo","Jutai do Igapo","Jutai Grande","Jutai Mirim","Jutai Peba","Jutai Pororoca","Jutai Roxo","Jutai-Acu","Jutany","Jutany de Campo","Jutay Branco","K'wannarri","Kakanjan Boesoe","Karvanari","Kawaknalli","Kawanari","Kwanari","Leather-Leaved Locust","Leathery-Leaved Locust","Legno Locusta","Locus","Locus Semirie","Locusboom","Locust","Locust Gum","Locustrier","Lokisi Kaka","Lokisie","Loksi","Lokus","Lokustbaum","Marbre","Masaicaran","Moire","Nazareno","Nere","Oleo de Jatai","Oleo Jatahy","Pacuy","Pakay","Palito","Palito Colorado","Pampa Estoraque","Pie de Venado","Pois Confiture","Quapinole","Quapinole Jutahy","Quauhpinolli","Quebra Machado","Rode Locus","Rode Rode Locus","Semirie","Simiri","Simiri Locust","Simirie","Sirari","South American Locust","Spruikhaanboom","Stinking Toe","Stinking-Toe","Surinaamsch Teakhout","Surixkra","Tema","Trapuca","Tsi-tsi-nau","Urapinol","West Indian Locust","Westindisch Teakhout","Westindsch Teakhout","Yatayba","Yutachi","Yutahi","Yutahy","Yutayba yutahy"], woodDb: "jatoba" },
  { name: "Brazilian Mahogany", setting: "0.47", scientific: "Swietenia macrophylla", synonyms: ["Acajou","Acajou Amerique","Acajou d'Amerique","Acajou du Honduras","Aguano","American Mahogany","Americkaans Mahonie","Amerikaans Mahonie","Aquano de Tabasco","Ara Pputange","Araputanga","Bastard lime","Bay-mahogany","Baywood","Belize mahogany","Big leafed mahogany","Big-leafed mahogany","Bigleaf mahogany","Broad leaved mahogany","Broad-leaved mahogany","Cabano","Caguano","Campeche","Caoba","Caoba Americana","Caoba de Atlantico","Caoba de Honduras","Caoba Hondurea","Caoba Hondurena","Caoba mahogany","Caoba roja","Caobilla","Cedro espinoso","Cedro-rana","Central American mahogany","Chacalte","Chiapas","Chiculte","Chiculti","Cobano","Costa Rica Mahogany","Costa Rico Mahogany","Crura","Cuban Mahogany","Flor de veradillo","Gateado","Giai Ngua","Granadillo","Guatemala Mahogany","Honduras Mahogany","Large Leaf Mahogany","Large Leaved Mahogany","Madeira","Mahogany","Mahogany Honduras","Mahoni","Mahonie","Mara","Mongo","Mongo do Rio Jurupari","Orura","Palo Xopliote","Palo Zopilote","Peruvian mahogany","Punab","Purab","Red Cedar","Red Wood","Resadillo","Sisam","Tabasco Mahogany","Tzopible","Tzopilote","Tzutzul","Venezuela Mahogany","Zopilocuahuitl","Zopilote","Zopilozontecomacuahuitl"], woodDb: "honduran-mahogany" },
  { name: "British Elm", setting: "0.53", scientific: "Ulmus procera", synonyms: ["English elm","Nave Elm","Red Elm","Vanlig Alm"], woodDb: "english-elm" },
  { name: "Bubinga", setting: "0.75", scientific: "Guibourtia spp.", synonyms: ["Kevasingo","Kevazingo","Waka"], woodDb: "bubinga" },
  { name: "Butternut", setting: "0.38", scientific: "Juglans cinerea", synonyms: ["Nogal","Nogal Blanco","Nogal Silvestre","Nuez Meca","Oilnut","Tocte","Tropical Walnut","Walnut","White Walnut"], woodDb: "butternut" },
  { name: "Cativo", setting: "0.42", scientific: "Prioria copaifera", synonyms: ["Algorrobillo","Amansamujer","Amanza Muger","Cabimbo","Cabimo","Cabino blanco","Camiba","Camibar","Canime","Cativo blanco","Cativo negro","Cautivo","Copachu","Copahyba","Copaiba","Copaibarana","Curucai","Eativo","Floresa","Kartiva","Muramo","Palo de Aceite","Red Cativo","Spanish Walnut","Tabasara","Tabosara","Taito","Timbo y Ata","Trementino","White Cativo"] },
  { name: "Cedar, Alaska", setting: "0.44", scientific: "Callitropsis nootkatensis", synonyms: ["Alaska Cypress","Alaska Ground Cypress","Alaska Yellow Cedar","Canoe cedar","False Cypress","Nootka Cypress","Nootka False Cypress","Nootka Sound Cypress","Pacific Coast Cypress","Sitka Cypress","White Cedar","Yellow Cypress"], woodDb: "alaskan-yellow-cedar" },
  { name: "Cedar, Atlantic White", setting: "0.32", scientific: "Chamaecyparis thyoides", synonyms: ["Chilopsis","Desert Willow","False Cypress","Southern White Cedar","Swamp Cedar","White Cedar"], woodDb: "atlantic-white-cedar" },
  { name: "Cedar, Eastern Red", setting: "0.47", scientific: "Juniperus virginiana", synonyms: ["Eastern redcedar","Juniper","Red Juniper","Redcedar","Savin","Virginia Pencil Cedar"], woodDb: "eastern-red-cedar" },
  { name: "Cedar, Incense", setting: "0.37", scientific: "Calocedrus decurrens", synonyms: ["Bastard cedar","California post cedar","Californian cedar","Californian incense cedar","Pecky Cedar","Pencil Cedar","Post Cedar","Red Cedar","White Cedar"], woodDb: "incense-cedar" },
  { name: "Cedar, Northern White", setting: "0.31", scientific: "Thuja occidentalis", synonyms: ["ArborV itae","Arborvitae","Cedar","Eastern Arborvitae","Eastern cedar","Eastern White Cedar","Eastern white-cedar","Northern Cedar","Northern White-Cedar","Swamp Cedar","White Cedar"], woodDb: "northern-white-cedar" },
  { name: "Cedar of Lebanon", setting: "0.53", scientific: "Cedrus libani", synonyms: ["Cedre du Liban","Cedro del Libano","True Cedar"], woodDb: "cedar-of-lebanon" },
  { name: "Cedar, Port Orford", setting: "0.43", scientific: "Chamaecyparis lawsoniana", synonyms: ["False Cypress","Ginuer Pine","Lawson Cypress","Lawson False Cypress","Lawson's Cypress","Oregon Cedar","Port Orford white cedar","White Cedar"], woodDb: "port-orford-cedar" },
  { name: "Cedar, Western Red", setting: "0.32", scientific: "Thuja plicata", synonyms: ["Arborvitae","British Columbia cedar","British Columbia red cedar","Canoe cedar","Giant Arbor-Vitae","Giant Arborvitae","Giant Cedar","Gigantic Cedar","Pacific Red Cedar","Pacific Redcedar","Red Cedar","Shinglewood","Western Arborvitae","Western Cedar","Western Redcedar"], woodDb: "western-red-cedar" },
  { name: "Cedar, Yellow", setting: "0.44", scientific: "Callitropsis nootkatensis", synonyms: ["Alaska Cypress","Alaska Ground Cypress","Alaska Yellow Cedar","Canoe cedar","False Cypress","Nootka Cypress","Nootka False Cypress","Nootka Sound Cypress","Pacific Coast Cypress","Sitka Cypress","White Cedar","Yellow Cypress"], woodDb: "alaskan-yellow-cedar" },
  { name: "Cedrella", setting: "0.39", scientific: "Cedrela odorata", synonyms: ["Acajou Rouge","Akuyari","Atoreb","Brazilian cedar","British Guiana cedar","British Honduras cedar","Cedar","Cedre rouge","Cedrela","Cedrela wood","Cedro","Cedro chino","Cedro hembra","Cedro macho","Cedro obscuro","Cedro oloroso","Cedro red","Cedro rojo","Central American cedar","Chujte","Cigar box cedar","Cigarbox cedar","Colorado cedro","Cuban cedar","Guyana cedar","Honduras cedar","Icte","Jamaican Cedar","Kalantas","Kapere","Koperi","Kurama","Kurana","Mexican Cedar","Nicaraguan Cedar","Parank","Paranka","Red Cedar","Rojas Cedar","Rosas Cedar","South American Cedar","Tabasco Cedar","Tiocuahuitl","Trinidad Cedar","Tropical Cedar","West Indian Cedar"], woodDb: "spanish-cedar" },
  { name: "Cherry, Black", setting: "0.50", scientific: "Prunus serotina", synonyms: ["American Black Cherry","Cabinet cherry","Capollin","Capuli","Capulin","Capulin cherry","Cerezo","Cerezo de Los Andes","Cherry","Chisos wild cherry","Choke cherry","Chokecherry","Detze","Edwards Plateau Cherry","Escarpment Cherry","Ghoto","Gila Chokecherry","Mountain Black Cherry","Muji","New England Mahogany","Pa-kshmuk","Plum","Rum Cherry","Southwestern Chokecherry","Spate Traubenkirsche","Tnunday","Whiskey Cherry","Wild Black Cherry","Wild Cherry","Xeugua"], woodDb: "black-cherry" },
  { name: "Chestnut, American", setting: "0.43", scientific: "Castanea dentata", synonyms: ["Chestnut","Chinkapin","English chestnut","European Chestnut","Spanish Chestnut"], woodDb: "american-chestnut" },
  { name: "Cocobolo", setting: "0.85", scientific: "Dalbergia retusa", synonyms: ["Jacarandaholz"], woodDb: "cocobolo" },
  { name: "Cottonwood, Balsam Poplar", setting: "0.34", scientific: "Populus balsamifera", synonyms: ["Balm","Balm Of Gilead","Balsam cottonwood","Balsam poplar","California poplar","Cottonwood","Hackmatack","Heartleaf Balsam Poplar","Poplar","Tacamahac","Tacamahac poplar","Western Balsam Poplar"], woodDb: "balsam-poplar" },
  { name: "Cottonwood, Black", setting: "0.35", scientific: "Populus trichocarpa", synonyms: ["Balm cottonwood","Cottonwood","Poplar","Western Balsam Poplar"], woodDb: "black-cottonwood" },
  { name: "Cottonwood, Eastern", setting: "0.40", scientific: "Populus deltoides", synonyms: ["Cottonwood","Deltoides-populier","Eastern Poplar","Necklace poplar","Poplar","Southern Cottonwood","Whitewood"], woodDb: "eastern-cottonwood" },
  { name: "Degame", setting: "0.72", scientific: "Calycophyllum candidissimum" },
  { name: "Determa", setting: "0.55", scientific: "Ocotea rubra" },
  { name: "Dogwood, Flowering", setting: "0.72", scientific: "Cornus florida", synonyms: ["Boxwood","Bunchberry","Cornel","Dogwood","Florida Dogwood"], woodDb: "dogwood" },
  { name: "Douglas Fir", setting: "0.48", scientific: "Pseudotsuga menziesii", synonyms: ["Blue Douglas-fir","British Columbia pine","British Columbian pine","Coast Douglas-fir","Colorado Douglas-fir","Colorado pino real","Colorado real","Columbian Pine","Douglas spruce","Douglas-fir","Douglas-fir (Coast)","Inland Douglas-Fir","Interior Douglas-Fir","Oregon Douglas-Fir","Oregon Pine","Puget Sound Pine","Red Fir","Rocky Mountain Douglas-Fir","Yellow Fir"], woodDb: "douglas-fir" },
  { name: "Ebony", setting: "0.94", scientific: "Diospyros spp.", synonyms: ["Abnus","Acha","African Ebony","Anang","Anang Gulod","Asian Black Ebony","Ata ata","Bale","Black ebony","Camagon","Dumbi","East indian ebony","Ebans","Ebone plaqueminier","Ebony persimmon","Indian Ebony","Itom Itom","Kakataki","Kamagong","Kanran","Karemara","Karimaran","Karu","Karunthali","Karunthoverai","Kaya Arang","Kayu Malam","Kendhu","Kukuo","Mallali","Marblewood","Mgiriti","Msindi","Mushtimbe","Nallati","Nalluti","Nyareti","Omenowa","Philippine Ebony","Shengutan","Tayung","Tendu","Trayung","Tuki","Tumbi","Tumiki","Ugau"], woodDb: "gaboon-ebony" },
  { name: "Elliotis Pine", setting: "0.59", scientific: "Pinus elliottii var. elliottii", synonyms: ["American Pitch Pine","Gulf Coast Pitch Pine","Pino Tea","Swamp Pine","Yellow Slash Pine"], woodDb: "slash-pine" },
  { name: "Elm, American", setting: "0.50", scientific: "Ulmus americana", synonyms: ["Florida Elm","Soft Elm","Swamp Elm","Water Elm","White Elm"], woodDb: "american-elm" },
  { name: "Elm, Rock", setting: "0.63", scientific: "Ulmus thomasii", synonyms: ["Canadian cork elm","Canadian rock elm","Cork elm","Hickory Elm"], woodDb: "rock-elm" },
  { name: "Elm, Slippery", setting: "0.53", scientific: "Ulmus rubra", synonyms: ["Gray Elm","Red Elm","Soft Elm"], woodDb: "red-elm" },
  { name: "English Cherry", setting: "0.58", scientific: "Prunus avium", synonyms: ["Cerisier","Cherry","Cherry Wood","European cherry","Fruit cherry","Gean","Kers","Kirsche","Mazzard","Merisier","Meurisier","Wild Cherry"], woodDb: "sweet-cherry" },
  { name: "English Oak", setting: "0.57", scientific: "Quercus robur", synonyms: ["Austrian Oak","Chene","Common English oak","Eiche","European oak","European white oak","French oak","Penduculate","Pendunculate Oak","Polish Oak","Quercia","Rovere","Skogsek","Slavonian Oak","Valhynian Oak","Yugoslavian oak"], woodDb: "english-oak" },
  { name: "European Ash", setting: "0.58", scientific: "Fraxinus excelsior", synonyms: ["Belgian ash","Common ash","English ash","Europeesche esche","French ash","Fresno","Hungarian ash","Italian Olive Ash","Olive Ash","Polish Ash","Slavonian Ash","Spanish Ash","Swedish Ash","Vanlig Ash","Vanlig Ask"], woodDb: "european-ash" },
  { name: "European Hornbeam", setting: "0.74", scientific: "Carpinus betulus", synonyms: ["Avenbok","Carpin","Carpy","Charme","Haagbeuk","Hagbuche","Hagebuche","Hainbuche","Hardbeam","Hornbaum","Hornbeam","Pine","Quickbeam","Quickenbeam","Quicktree","Vitbok","Weissbuche","Yoke-Elm"], woodDb: "european-hornbeam" },
  { name: "European Walnut", setting: "0.56", scientific: "Juglans regia", synonyms: ["Walnut"], woodDb: "english-walnut" },
  { name: "Fir, Balsam", setting: "0.35", scientific: "Abies balsamea", synonyms: ["Balsam","Blister fir","Bracted balsam fir","Canadian balsam","Canadian fir","Eastern Fir","Galm of Gilead Fir","Silver Pine"], woodDb: "balsam-fir" },
  { name: "Fir, California Red", setting: "0.38", scientific: "Abies magnifica", synonyms: ["Golden Fir","Red Fir","Shasta Fir","Shasta Red Fir","Silvertip","Western Fir"], woodDb: "california-red-fir" },
  { name: "Fir, Grand", setting: "0.37", scientific: "Abies grandis", synonyms: ["Epicea","Giant Fir","Larch","Lowland Fir","Lowland White Fir","Menzies Fir","Oregon Fir","Sapin","Silver Fir","Western Balsam Fir","Western Fir","Western White Fir","Yellow Fir"], woodDb: "grand-fir" },
  { name: "Fir, Noble", setting: "0.39", scientific: "Abies procera", synonyms: ["Oregon Larch","Red Fir","Western Fir"], woodDb: "noble-fir" },
  { name: "Fir, Pacific Silver", setting: "0.43", scientific: "Abies amabilis", synonyms: ["Alpine Fir","Amabilis Fir","Balsam","Cascade fir","Great Silver Fir","Larch","Lovely Fir","Red Fir","Silver Fir","Western Fir"], woodDb: "pacific-silver-fir" },
  { name: "Fir, Subalpine", setting: "0.32", scientific: "Abies lasiocarpa", synonyms: ["Alpine Fir","Balsam","Corkbark fir","Pino Real Blanco","Rocky Mountain Fir","Western Balsam Fir","Western Fir","White Balsam"], woodDb: "subalpine-fir" },
  { name: "Fir, White", setting: "0.39", scientific: "Abies concolor", synonyms: ["Colorado fir","Colorado white fir","Concolor fir","Lows Fir","Oyamel","Pacific White Fir","Pino Real Blanco","Rocky Mountain White Fir","Silver Fir","Western Fir","White Balsam"], woodDb: "white-fir" },
  { name: "Gombeira", setting: "1.00", scientific: "Didelotia africana" },
  { name: "Guatambu (Argentinean)", setting: "0.70", scientific: "Balfourodendron riedelianum" },
  { name: "Guatambu (Brazil)", setting: "0.79", scientific: "Aspidosperma spp." },
  { name: "Gum, Black", setting: "0.50", scientific: "Nyssa sylvatica", synonyms: ["Chan thip","Lau Tau","Mascalwood","Pepperidge","Resak","Sourgum","Taungsagaing","Tupelo","Tupelo Gum"], woodDb: "black-tupelo" },
  { name: "Gum, Red", setting: "0.52", scientific: "Liquidambar styraciflua", synonyms: ["Alligator Tree","American Red Gum","American Styrax","Balsamo blanco","Bilstead","Bilsted","Blistead","Blisted","Copalillo","Copalone","Diquidambo","Estoraque","Gum Wood","Hazel Pine","Hazel Wood","Icob","Ien-gau-o","Ingano","Ko'ma","Ko'ma'liso","Liquidambar","Liquidambo","Mola","Nijte-Pijto","Nite-biito","Ocozote","Quivambaro","Sap Gum","Satin Walnut","Skchute","Slu'to'nko","Somerio","Starleaf Gum","Storax","Sweet Gum","Xochicatscahuitl","Yaga-Bizigui","Yaga-Huille","Yellow Gum"], woodDb: "sweetgum" },
  { name: "Hackberry", setting: "0.53", scientific: "Celtis occidentalis", synonyms: ["Bastard elm","Common hackberry","False Elm","Hacktree","Hoop Ash","Nettle Tree","Nettletree","Sugarberry","Western hackberry"], woodDb: "hackberry" },
  { name: "Hemlock, Eastern", setting: "0.40", scientific: "Tsuga canadensis", synonyms: ["American hemlock","Canada hemlock","Canadian hemlock","Common hemlock","Hemlock","Hemlock spruce","White Hemlock"], woodDb: "eastern-hemlock" },
  { name: "Hemlock, Mountain", setting: "0.45", scientific: "Tsuga mertensiana", woodDb: "mountain-hemlock" },
  { name: "Hemlock, Western", setting: "0.45", scientific: "Tsuga heterophylla", synonyms: ["Alaska Pine","British Columbia hemlock","British Columbian hemlock","Gray Fir","Grey Fir","Hemlock spruce","Huron Pine","Pacific Coast Hemlock","Pacific Hemlock","Prince Albert Fir","Prince Albert Spruce","Silver Fir","Western Hemlock Fir","Western Hemlock Spruce","White Hemlock"], woodDb: "western-hemlock" },
  { name: "Hickory (Pecan), Bitternut", setting: "0.66", scientific: "Carya cordiformis", synonyms: ["Bitternut","Bitternut hickory","Hickory","Swamp Hickory"], woodDb: "bitternut-hickory" },
  { name: "Hickory (Pecan), Nutmeg", setting: "0.60", scientific: "Carya myristiciformis", synonyms: ["Hickory","Nutmeg Hickory"], woodDb: "nutmeg-hickory" },
  { name: "Hickory (Pecan), Water", setting: "0.62", scientific: "Carya aquatica", synonyms: ["Bitter pecan","Bitter water hickory","Hickory","Swamp Hickory","Water Hickory","Wild Pecan"], woodDb: "water-hickory" },
  { name: "Hickory (True), Mockernut", setting: "0.72", scientific: "Carya tomentosa", synonyms: ["Hickory","Mockernut","Mockernut Hickory","White Hickory"], woodDb: "mockernut-hickory" },
  { name: "Hickory (True), Pignut", setting: "0.75", scientific: "Carya glabra", synonyms: ["Pignut Hickory"], woodDb: "pignut-hickory" },
  { name: "Hickory (True), Shagbark", setting: "0.72", scientific: "Carya ovata", synonyms: ["Hickory","Scalybark Hickory","Shagbark Hickory","Shellbark Hickory"], woodDb: "shagbark-hickory" },
  { name: "Hickory (True), Shellbark", setting: "0.69", scientific: "Carya laciniosa", synonyms: ["Big Shellbark Hickory","Kingnut","Kingnut Hickory"], woodDb: "shellbark-hickory" },
  { name: "Hickory, Pecan", setting: "0.66", scientific: "Carya illinoinensis", synonyms: ["Bitter pecan","Hickory","Nogal Morado","Nuez Encarcelada","Sweet Pecan"], woodDb: "pecan" },
  { name: "Holly, American", setting: "0.55", scientific: "Ilex opaca", synonyms: ["Dune Holly","Holly","Hummock Holly","Scrub Holly","White Holly"], woodDb: "holly" },
  { name: "Hophornbeam, Eastern", setting: "0.70", scientific: "Ostrya virginiana", synonyms: ["Ironwood"], woodDb: "hophornbeam" },
  { name: "Hura", setting: "0.40", scientific: "Hura crepitans", synonyms: ["Acacu","Acau","Acaupa","Acaupar","Arbol del Ddiablo","Arbre au Diable","Areeiro","Arenillero","Arerillo","Asniakara","Assacu","Bois du diable","Castana","Castaneto","Catahua","Cataua","Ceiba amarilla","Ceiba bla","Ceiba blanca","Ceiba de leche","Ceiba habilla","Ceiba habillo","Ceiba mil pesos","Cuatatachi","Haba","Haba de Indio","Habillo","Hura Wood","Jabillo","Javillo","Molinillo","Monkeys Dinner Bell","Nune","Ochoho","Ochoo","Ovillo","Pet du Diable","Possentrie","Possum","Possumwood","Quauhayohuatli","Quauhtlat-Latzin","Racuada","Rakuda","Rakudar","Sablier","Salvadera","Sandbox","Seda Blanca","Solimanche","Tetereta","Tronador","Uacacu"] },
  { name: "Indian Laurel", setting: "0.79", scientific: "Terminalia tomentosa", woodDb: "indian-laurel" },
  { name: "Ipe", setting: "0.99", scientific: "Handroanthus spp.", synonyms: ["Acapro","Akkeja","Akkekeja","Ala-onni","Ala-ore","Alahorre","Alan-che","Alcapro","Alumbre","Amapa","Amapa Prieta","Amapa Priete","Amapa Prieto","Aoka","Arabore","Araguaney","Arahoni","Aravaney","Arawnig","Arawnig-yek","Arco","Arcwood","Arowore","Arra-ore","Arrhonee","Aruain","Bastard lignum vitae","Bethabara","Bois d'ebene verte","Bois d'evilasse","Bow wood","Bow-wood","Brazilian Walnut","Caexeta","Canada","Canaguate","Canahuate","Chicala","Cogwood","Coralibe","Cortes","Cortes amarillo","Cortez","Cortez amarillo","Cortez colorado","Cortez de venado","Corteza","Courali","Curari","Curarire","Ebano verde","Ebene soufre","Ebene vert","Ebene verte","Echahumo","Enbotta-koenatjepre","Flor amarillo","Greenheart","Grenhatti","Grienharti","Groenhart","Groenhati","Guayacan","Guayacan polvillo","Gupariba","Hackia","Hackoyia","Haekia","Hahuache","Hakia","Hakkea","Ijzerhout","Ipe amarillo","Ipe cascudo","Ipe de varzea","Ipe do compo","Ipe folhas roxas","Ipe jabotica","Ipe preto","Ipe roxo","Ipe tabaco","Ipe una","Irontree","Ironwood","Konawadranup","Lapachillo Tally","Lapacho","Lapacho Amarillo","Lapacho Blanco","Lapacho Crespo","Lapacho Negro","Lapacho Rosa","Lubre","Madera Negra","Makagrien","Makka Groenhart","Mangienhatti","Mano de Leon","Masicaran","Noibwood","Pao d'Arco","Pao d'Arco Amarillo","Pao d'Arco Roxo","Pau d'Arco","Pau Darco","Penda","Polvillo","Poui","Quebracho","Quiarapaiba","Ranoi","Roble Cinero","Surinam Greenheart","Tahua","Tahuari","Tamura Tuira","Tauary","Taye","Tayi","Urupariba","Verdecillo","Washiba","Wasiba","Wasieba","Wassiba","Wehete","Whoua-Whoua","Woile","Xha-hua-che","Yellow Guayacan","Yellow Poui"], woodDb: "ipe" },
  { name: "Iroko", setting: "0.57", scientific: "Milicia excelsa", synonyms: ["Abang","African Oak","African Teak","Agui","Akede","Bakana","Bang","Banghi","Bobang","Bonzo","Bush oak","Bush teak","Bwagashanga","Cambala","Corkwood","Elowa","Elua","Elui","Emang","Guele","Gutumba","Intule","Kambala","Kimurumba","Logo Asagu","Loko","Lusanga","Mamangi","Mandji","Mbala","Mbara","Mereira","Mgunda","Minarui","Mokongo","Moloundou","Molundu","Moreira","Muberry","Mucoco","Murie","Murumba","Mururi","Mutumba","Mutumbav","Muvule","Mvule","Mvuli","Myule","Nsan","Ntong","Obas Tree","Odji","Odoum","Odum","Oduna","Ofryio","Olia","Olua","Olwaa","Oroko","Oroko Ulokoodigpe","Roco","Rokko","Sanga","Semei","Semli","Sili","Sime","Simli","Simme","Ssare","Tema","Timmi","Tomboiro Noir","Toumbohiro Noir","Tule","Tule Mufala","Uklobce","Ulok","Uloko-Mushinogbon","Ulundu"], woodDb: "iroko" },
  { name: "Jacaranda", setting: "0.34", scientific: "Jacaranda mimosifolia", synonyms: ["Blue Jacaranda"] },
  { name: "Jarrah", setting: "0.75", scientific: "Eucalyptus marginata", woodDb: "jarrah" },
  { name: "Jelutong", setting: "0.38", scientific: "Dyera costulata", woodDb: "jelutong" },
  { name: "Kapur", setting: "0.70", scientific: "Dryobalanops spp." },
  { name: "Karri", setting: "0.79", scientific: "Eucalyptus diversicolor", woodDb: "karri" },
  { name: "Keruing", setting: "0.76", scientific: "Dipterocarpus spp.", woodDb: "keruing" },
  { name: "Kingwood", setting: "1.16", scientific: "Dalbergia cearensis", woodDb: "kingwood" },
  { name: "KOA (Acacia Koa)", setting: "0.63", scientific: "Acacia koa", woodDb: "koa" },
  { name: "Larch, Euro", setting: "0.48", scientific: "Larix decidua", synonyms: ["Common larch","European larch","Lark"], woodDb: "european-larch" },
  { name: "Larch, Western", setting: "0.52", scientific: "Larix occidentalis", synonyms: ["Hackmatack","Hackmatack larch","Larch","Montana Larch","Mountain Larch","Western Tamarack"], woodDb: "western-larch" },
  { name: "Laurel, California", setting: "0.55", scientific: "Umbellularia californica", synonyms: ["Bay laurel","Myrtle","Pepperwood","Spice Tree"], woodDb: "oregon-myrtle" },
  { name: "Lignum Vitae", setting: "1.13", scientific: "Guaiacum officinale", woodDb: "lignum-vitae" },
  { name: "Limba", setting: "0.40", scientific: "Terminalia superba", synonyms: ["Afara","Afia Afia","Afodonko","Afraa","Aghan","Akam","Akom","An Rin","Bale","Baya","Bese","Blie","Bokone","Chene limbo","Chene-limbo","Congo walnut","Dark Limba","Dark Noir","Djombe","Egean","Egoin nufua","Egonni","Egoyin","Faraen","Frake","Fram","Frameri","Framo","Frane","Frango","Frany","Gbararada","Ka-Ren","Kegblale","Kojaagei","Kojagei","Kone","Kongo","Korina","Kosina","Kumkunbe","Landi","Light Limba","Limba Clair","Limba Noir","Limbo","Moukonia","Mukonja","N'dimba","N'kom","N'limba","Noyer","Noyer du Mayombe","Noyer Limbo","Offram","Ofram","Ojiloko","Owebala","Shingle Wood","Unwonrom","Weiss","White Afara","White Limba","Yellow pine"], woodDb: "limba" },
  { name: "Locust, Black", setting: "0.69", scientific: "Robinia pseudoacacia", synonyms: ["Acacia","False Acacia","Golden Oak","Green Locust","Loco","Locust","Post Locust","Red Locust","Robinia","Shipmast Locust","Virginische Schotendorn","White Locust","Yellow Locust"], woodDb: "black-locust" },
  { name: "Macassar Ebony", setting: "0.90", scientific: "Diospyros celebica", synonyms: ["Asian Grained Ebony","Calamander wood","Camagon","Coromandel","Ebene veinee d'Asie","Golden ebony","Indian Ebony","Temru","Tendu","Timbruni","Tunki"], woodDb: "macassar-ebony" },
  { name: "Madrone, Pacific", setting: "0.64", scientific: "Arbutus menziesii", synonyms: ["Arbuti Tree","Coast madrone","Madrona","Madrone","Madrono"], woodDb: "madrone" },
  { name: "Magnolia, Southern", setting: "0.50", scientific: "Magnolia grandiflora", synonyms: ["Bat tree","Big Laurel","Black lin","Bullbay or Bull bay","Cucumber Wood","Evergreen Magnoilia","Magnolia","Mountain Magnolia","Sweet Magnolia"], woodDb: "southern-magnolia" },
  { name: "Mahogany, African", setting: "0.44", scientific: "Khaya spp.", synonyms: ["Acajou Umbaua","Banket mahogany","East African mahogany","Iluli","Kaonde","M'bane","Mahogany","Mbamba","Mbaua","Mbawa","Mkangazi","Mozambique Mahogany","Mtondoo","Mtondoro","Mubaba","Mubawa","Mululu","Mururu","Muvava","Muwawa","Myofu","Nyasaland Mahogany","Red Mahogany","Umbaba","Umbaua"], woodDb: "african-mahogany" },
  { name: "Mahogany, True", setting: "0.47", scientific: "Swietenia macrophylla", synonyms: ["Acajou","Acajou Amerique","Acajou d'Amerique","Acajou du Honduras","Aguano","American Mahogany","Americkaans Mahonie","Amerikaans Mahonie","Aquano de Tabasco","Ara Pputange","Araputanga","Bastard lime","Bay-mahogany","Baywood","Belize mahogany","Big leafed mahogany","Big-leafed mahogany","Bigleaf mahogany","Broad leaved mahogany","Broad-leaved mahogany","Cabano","Caguano","Campeche","Caoba","Caoba Americana","Caoba de Atlantico","Caoba de Honduras","Caoba Hondurea","Caoba Hondurena","Caoba mahogany","Caoba roja","Caobilla","Cedro espinoso","Cedro-rana","Central American mahogany","Chacalte","Chiapas","Chiculte","Chiculti","Cobano","Costa Rica Mahogany","Costa Rico Mahogany","Crura","Cuban Mahogany","Flor de veradillo","Gateado","Giai Ngua","Granadillo","Guatemala Mahogany","Honduras Mahogany","Large Leaf Mahogany","Large Leaved Mahogany","Madeira","Mahogany","Mahogany Honduras","Mahoni","Mahonie","Mara","Mongo","Mongo do Rio Jurupari","Orura","Palo Xopliote","Palo Zopilote","Peruvian mahogany","Punab","Purab","Red Cedar","Red Wood","Resadillo","Sisam","Tabasco Mahogany","Tzopible","Tzopilote","Tzutzul","Venezuela Mahogany","Zopilocuahuitl","Zopilote","Zopilozontecomacuahuitl"], woodDb: "honduran-mahogany" },
  { name: "Manni", setting: "0.63", scientific: "Symphonia globulifera" },
  { name: "Maple, Bigleaf", setting: "0.48", scientific: "Acer macrophyllum", synonyms: ["Broadleaf maple","Maple","Oregon Maple","Pacific Coast Maple","Western Maple"], woodDb: "bigleaf-maple" },
  { name: "Maple, Black", setting: "0.57", scientific: "Acer nigrum", synonyms: ["Black sugar maple","Hard Rock Maple","Maple","Rock Maple"], woodDb: "black-maple" },
  { name: "Maple, Hard", setting: "0.60", scientific: "Acer saccharum", synonyms: ["Bird's eye maple","Blister maple","Canadian maple","Curly Maple","Fiddleback Maple","Maple","Rock Maple"], woodDb: "hard-maple" },
  { name: "Maple, Red", setting: "0.54", scientific: "Acer rubrum", synonyms: ["Carolina red maple","Drummond red maple","Maple","Scarlet Maple","Swamp Maple","Water Maple","White Maple"], woodDb: "red-maple" },
  { name: "Maple, Silver", setting: "0.47", scientific: "Acer saccharinum", synonyms: ["White Maple"], woodDb: "silver-maple" },
  { name: "Maple, Soft", setting: "0.49", scientific: "Acer rubrum", synonyms: ["Carolina red maple","Drummond red maple","Maple","Scarlet Maple","Swamp Maple","Water Maple","White Maple"], woodDb: "red-maple" },
  { name: "Maple, Sugar", setting: "0.63", scientific: "Acer saccharum", synonyms: ["Bird's eye maple","Blister maple","Canadian maple","Curly Maple","Fiddleback Maple","Maple","Rock Maple"], woodDb: "hard-maple" },
  { name: "Merbau", setting: "0.67", scientific: "Intsia spp.", woodDb: "merbau" },
  { name: "Mersawa", setting: "0.54", scientific: "Anisoptera spp." },
  { name: "Mesquite", setting: "0.86", scientific: "Prosopis spp.", woodDb: "honey-mesquite" },
  { name: "Monkeypod", setting: "0.50", scientific: "Samanea saman", synonyms: ["Monkey Pod"], woodDb: "monkeypod" },
  { name: "Mountain Ash (Eucalyptus)", setting: "0.62", scientific: "Eucalyptus regnans", synonyms: ["Tasmanian Oak"], woodDb: "mountain-ash" },
  { name: "Movingui", setting: "0.72", scientific: "Distemonanthus benthamianus", synonyms: ["Ayan","Ayanran","Nigerian Satinwood"], woodDb: "movingui" },
  { name: "Muninga", setting: "0.59", scientific: "Pterocarpus angolensis", synonyms: ["Ambila","Bloodwood","Brown African padauk","Girassonde","Imbilo","Kajaat","Kajat","Kajatenhout","Kejaat","Kiaat","Kiatt","Maninga","Mavamaropa","Mbila","Mlombwa","Mninga","Moroto","Mtumbati","Mukurambira","Mukwa","Mulambi","Mulombe","Mulombwa","Munhaneca","Mutete","Mututi","Muzwamaloa","N'dombe","Rhodesian Walnut","Sealing Wax Tree","Thondo","Thondu","Thordo","Transvaal Kajatenhout","Umbila","Umvagaz","Umvangazi"], woodDb: "muninga" },
  { name: "Myrtle, Oregon", setting: "0.55", scientific: "Umbellularia californica", synonyms: ["Bay laurel","Myrtle","Pepperwood","Spice Tree"], woodDb: "oregon-myrtle" },
  { name: "Myrtle, Tasmanian", setting: "0.64", scientific: "Nothofagus cunninghamii", woodDb: "tasmanian-myrtle" },
  { name: "Oak (Red), Black", setting: "0.61", scientific: "Quercus velutina", synonyms: ["Black Oak","Cucharillo","Encino","Encino negro","Mamecillo","Quercitron","Quercitron oak","Red Oak","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Smooth Bark Oak","Yellow Bark Oak","Yellow Oak"], woodDb: "black-oak" },
  { name: "Oak (Red), Cherrybark", setting: "0.68", scientific: "Quercus pagoda", synonyms: ["Cherry bark oak"], woodDb: "cherrybark-oak" },
  { name: "Oak (Red), Laurel", setting: "0.63", scientific: "Quercus laurifolia", synonyms: ["Laurel Oak"], woodDb: "laurel-oak" },
  { name: "Oak (Red), Northern", setting: "0.63", scientific: "Quercus rubra", synonyms: ["Black oak","Canadian red oak","Gray Oak","Northern Red Oak","Red Oak","Southern Red Oak"], woodDb: "red-oak" },
  { name: "Oak (Red), Pin", setting: "0.63", scientific: "Quercus palustris", synonyms: ["Cucharillo","Encino","Encino negro","Mamecillo","Pin Oak","Red Oak","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Spanish Oak","Spanish Swamp Oak","Swamp Oak","Water Oak"], woodDb: "pin-oak" },
  { name: "Oak (Red), Scarlet", setting: "0.67", scientific: "Quercus coccinea", synonyms: ["Black oak","Cucharillo","Encino","Encino negro","Mamecillo","Red Oak","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Scarlet Oak","Spanish Oak"], woodDb: "scarlet-oak" },
  { name: "Oak (Red), Southern", setting: "0.59", scientific: "Quercus falcata", synonyms: ["Southern Red Oak"], woodDb: "southern-red-oak" },
  { name: "Oak (Red), Water", setting: "0.63", scientific: "Quercus nigra", synonyms: ["Water Oak"], woodDb: "water-oak" },
  { name: "Oak (Red), Willow", setting: "0.69", scientific: "Quercus phellos", synonyms: ["Willow Oak"], woodDb: "willow-oak" },
  { name: "Oak (White), Bur", setting: "0.64", scientific: "Quercus macrocarpa", synonyms: ["Blue oak","Bur Oak","Cucharillo","Encino","Encino negro","Mamecillo","Mossy Overcup Oak","Mossycup Oak","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Scrub Oak"], woodDb: "bur-oak" },
  { name: "Oak (White), Chestnut", setting: "0.66", scientific: "Quercus prinus", synonyms: ["Chestnut white oak"], woodDb: "chestnut-oak" },
  { name: "Oak (White), Overcup", setting: "0.63", scientific: "Quercus lyrata", synonyms: ["Overcup White Oak"], woodDb: "overcup-oak" },
  { name: "Oak (White), Post", setting: "0.67", scientific: "Quercus stellata", synonyms: ["Cucharillo","Encino","Encino negro","Mamecillo","Post Oak","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito"], woodDb: "post-oak" },
  { name: "Oak (White), Swamp", setting: "0.72", scientific: "Quercus bicolor", synonyms: ["Cucharillo","Encino","Encino negro","Mamecillo","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Swamp White Oak"], woodDb: "swamp-white-oak" },
  { name: "Oak (White), Swamp Chestnut", setting: "0.67", scientific: "Quercus michauxii", synonyms: ["Cow oak","Cucharillo","Encino","Encino negro","Mamecillo","Roble #ncino","Roble Amarillo","Roble Colorado","Roblecito","Swamp Chestnut Oak"], woodDb: "swamp-chestnut-oak" },
  { name: "Oak, California Black", setting: "0.53", scientific: "Quercus kelloggii", synonyms: ["Black oak","Kellogg Oak"], woodDb: "california-black-oak" },
  { name: "Oak, White", setting: "0.68", scientific: "Quercus alba", synonyms: ["Arizona Oak","Arizona White Oak","Cucharillo","Encino","Encino negro","Mamecillo","Roble Amarillo","Roble Colorado","Roble Encino","Roblecito","Stave Oak"], woodDb: "white-oak" },
  { name: "Obeche", setting: "0.32", scientific: "Triplochiton scleroxylon", synonyms: ["Abachi","African Bush Maple","African Maple","African Whitewood","Ajuss","Arare","Arere","Ayos","Ayous","Ayus","Bado","Bamba","Batobus","Bush maple","Cofa","Egin-fifen","Ejoung","Ejuong","Ewowo","Ghana Obeche","Hafa","Hofa","Kofa","Larana Whitewood","Lomangene","M'Bado","Nkom","Obechi","Obeke","Okpa","Okpo","Otrotso","Oua-Oua","Ouesse","Owawa","Owowa","Pataboa","Sama","Samba","Samba Gris","Samba Ou Ayous","Samba-Ayous","Sankamba","Satinwood","Serama","Soft Satinwood","Wana","Wawa","Wawa Arera"], woodDb: "obeche" },
  { name: "Okoume", setting: "0.35", scientific: "Aucoumea klaineana", woodDb: "okoume" },
  { name: "Olive", setting: "0.81", scientific: "Olea europaea", synonyms: ["Common Olive","European Olive"], woodDb: "olive" },
  { name: "Opepe", setting: "0.68", scientific: "Nauclea diderrichii", woodDb: "opepe" },
  { name: "Padauk (P. indicus)", setting: "0.57", scientific: "Pterocarpus indicus", synonyms: ["Amboyna","Andaman Padduik","Andaman Redwood","Angsama","Angsana","Chalanga-da","East Indian mahogany","Indian redwood","Liki","Nara","Narra","Narravitail","New Guinea Rosewood","Nonalu","Papua New Guniea Rosewood","Red Narra","Ringii","Rosewood","Sena","Solomons Padauk","Sonokembang","Vermilion Wood","Warave","Yaya Sa","Yellow Narra"], woodDb: "narra" },
  { name: "Padauk (P. macrocarpus)", setting: "0.79", scientific: "Pterocarpus macrocarpus", synonyms: ["Burma padauk","Mai Pradoo","Mai-Chi-Tawk","Mai-Pi-Tawk","Pradoo","Pterocarpus"], woodDb: "burma-padauk" },
  { name: "Padauk (P. marsupium)", setting: "0.71", scientific: "Pterocarpus marsupium", synonyms: ["Angu","Bebe","Bijasal","Huevos de Gato","Indian Padauk","Lagunero","Mututi","Nogal Falso","Palo de Pollo","Pau Sangua","Sangre","Sangre de Drago","Sangrillo","Yaya Sangre"] },
  { name: "Parana Pine", setting: "0.49", scientific: "Araucaria angustifolia", synonyms: ["Araucaria","Brazilian araucaria","Brazilian pine","Chilean pine","Curiy","Cury","Inho Vermelho","Kuviy","Monkey Puzzle Tree","Pehuen","Pilon","Pinheiro","Pinheiro do Brasil","Pinheiro do Parana","Pinho","Pinho Branco","Pinho Brasileiro","Pinho Brasilero","Pinho do Parana","Pino","Pino Blanco","Pino Paranâ€¡"], woodDb: "parana-pine" },
  { name: "Pecan", setting: "0.60", scientific: "Carya illinoinensis", synonyms: ["Bitter pecan","Hickory","Nogal Morado","Nuez Encarcelada","Sweet Pecan"], woodDb: "pecan" },
  { name: "Peroba de Campos", setting: "0.66", scientific: "Paratecoma peroba" },
  { name: "Peroba Rosa", setting: "0.71", scientific: "Aspidosperma peroba", woodDb: "peroba-rosa" },
  { name: "Persimmon, Common", setting: "0.71", scientific: "Diospyros virginiana", synonyms: ["American Ebony","Bara-bara","Boa-wood","Butterwood","Cylil Date Plum","Persimmon","Possum Wood","Virginia Date Palm","White Ebony"], woodDb: "persimmon" },
  { name: "Pine, Eastern White", setting: "0.35", scientific: "Pinus strobus", synonyms: ["American Yellow Pine","Austrian White Pine","Canadain white pine","Canadian yellow pine","Cork pine","Northern White Pine","Ottawa Pine","Ottawa White Pine","Pattern Pine","Pumpkin Pine","Quebec Pine","Quebec Yellow Pine","Sapling Pine","Soft Pine","Weymouth Pine","Yellow Pine"], woodDb: "eastern-white-pine" },
  { name: "Pine, Hoop", setting: "0.44", scientific: "Araucaria cunninghamii", synonyms: ["Australian Araucaria","Bunya bunya","Norfolk Island Pine","Pin Colonnaire","Sapin de Montagne"], woodDb: "hoop-pine" },
  { name: "Pine, Jack", setting: "0.43", scientific: "Pinus banksiana", synonyms: ["Banks pine","Banksian pine","Black pine","Cypress Pine","Gray Pine","Grey Pine","Hudson Bay pine","Juniper Bull Pine","Northern Scrub Pine","Pine","Princess Pine","Scrub Pine"], woodDb: "jack-pine" },
  { name: "Pine, Loblolly", setting: "0.51", scientific: "Pinus taeda", synonyms: ["Bassett pine","Foxtail Pine","Indian Pine","North Carolina pine","Oldfield Pine","Pine","Pinho-Teda","Swamp Pine","Taeda Pine","Torch Pine","Yellow Pine"], woodDb: "loblolly-pine" },
  { name: "Pine, Lodgepole", setting: "0.41", scientific: "Pinus contorta", synonyms: ["Black pine","Contorta pine","Knotty Pine","Scrub Pine","Shore Pine","Tamarack Pine","Western Jack Pine"], woodDb: "lodgepole-pine" },
  { name: "Pine, Longleaf", setting: "0.59", scientific: "Pinus palustris", synonyms: ["American Pitch Pine","Cedar pine","Florida Longleaf Pine","Florida Yellow Pine","Georgia Pine","Georgia Yellow Ppine","Hard Pine","Heart Pine","Long Leaf Pitch Pine","Longleaf Pitch Pine","Longstraw Pine","Palustris Pine","Palustris-Den","Pine","Straw Pine","Texas Yellow Pine","Walter Pine","Yellow Ppine"], woodDb: "longleaf-pine" },
  { name: "Pine, Pitch", setting: "0.52", scientific: "Pinus rigida", synonyms: ["Pine"], woodDb: "pitch-pine" },
  { name: "Pine, Pond", setting: "0.56", scientific: "Pinus serotina", synonyms: ["Marsh Pine","Pine","Pocosin pine"], woodDb: "pond-pine" },
  { name: "Pine, Ponderosa", setting: "0.40", scientific: "Pinus ponderosa", synonyms: ["Big pine","Bird's-eye pine","Black jack pine","Black pine","Blackjack pine","British Colombia soft pine","Bull pine","Californian white pine","Knotty Pine","Oregon Pine","Pine","Pole Pine","Pondosa Pine","Prickly Pine","Western Pine","Western Soft Pine","Western Yellow Pine","Yellow Pine"], woodDb: "ponderosa-pine" },
  { name: "Pine, Red", setting: "0.46", scientific: "Pinus resinosa", synonyms: ["American Red Pine","Canadian red pine","Hard Pine","Norway Pine","Ottawa Red Pine","Pig Iron Pine","Pine","Quebec Red Pine","Shellbark Norway"], woodDb: "red-pine" },
  { name: "Pine, Sand", setting: "0.48", scientific: "Pinus clausa", synonyms: ["Scrub Pine"], woodDb: "sand-pine" },
  { name: "Pine, Scots", setting: "0.45", scientific: "Pinus sylvestris", synonyms: ["Archangel Redwood","Baltic fir","Baltic pine","Baltic redwood","Common pine","Danzig fir","Danzig pine","Finnish fir","Finnish redwood","Gefle fir","Memel Fir","Norway Fir","Pine","Polish Redwood","Red Deal","Redwood","Scotch Pine","Scots fir","Siberian Redwood","Soderhamn Fir","Swedish Fir","Swedish Redwood","Vanlig Tall","Vanligtall","White Sea Fir","Yellow Deal"], woodDb: "scots-pine" },
  { name: "Pine, Shortleaf", setting: "0.51", scientific: "Pinus echinata", synonyms: ["Hard Pine","Shortstraw Pine"], woodDb: "shortleaf-pine" },
  { name: "Pine, Slash", setting: "0.59", scientific: "Pinus elliottii", synonyms: ["American Pitch Pine","Gulf Coast Pitch Pine","Pino Tea","Swamp Pine","Yellow Slash Pine"], woodDb: "slash-pine" },
  { name: "Pine, Spruce", setting: "0.44", scientific: "Pinus glabra", synonyms: ["Cedar pine","Pine","Walter Pine"], woodDb: "spruce-pine" },
  { name: "Pine, Sugar", setting: "0.36", scientific: "Pinus lambertiana", synonyms: ["Big pine","Californian sugar pine","Californina soft pine","Gigantic Pine","Great Sugar Pine","Pine","Shade Pine"], woodDb: "sugar-pine" },
  { name: "Pine, Virginia", setting: "0.48", scientific: "Pinus virginiana", synonyms: ["Jersey Pine","Pine","Scrub Pine","Southern Pine"], woodDb: "virginia-pine" },
  { name: "Pine, Western White", setting: "0.35", scientific: "Pinus monticola", synonyms: ["Idaho White Pine","Mountain Pine","Mountain White Pine","Pine","Silver Pine","Soft Pine","White Pine"], woodDb: "western-white-pine" },
  { name: "Plane (Lacewood)", setting: "0.49", scientific: "Platanus spp.", synonyms: ["English plane","European plane","French plane","London Plane","Platane"], woodDb: "london-plane" },
  { name: "Poplar, Yellow", setting: "0.42", scientific: "Liriodendron tulipifera", synonyms: ["American Tulipwood","American Whitewood","Blue poplar","Canadian whitewood","Canary whitewood","Canary wood","Canoe wood","Green Cypress","Hickory Poplar","Poplar","Popple","Saddle Tree","Saddletree","Tulip Poplar","Tuliptree","Tulipwood","Virginian Poplar","White Poplar","Whitewood","Yellow-Wood"], woodDb: "yellow-poplar" },
  { name: "Primavera", setting: "0.42", scientific: "Cybistax donnell-smithii", woodDb: "primavera" },
  { name: "Purpleheart", setting: "0.71", scientific: "Peltogyne spp.", woodDb: "purpleheart" },
  { name: "Radiata Pine", setting: "0.45", scientific: "Pinus radiata", synonyms: ["Insignis","Insignis Pine","Insignis-Den","Insular Pine","Monterey Pine","Pino Iinsigne","Remarkable Pine"], woodDb: "radiata-pine" },
  { name: "Ramin", setting: "0.56", scientific: "Gonystylus spp.", synonyms: ["Ainunura","Fungunigalo","Garu-buaja","Lanutan-Bagio","Lanutan-Bagyo","Latareko","Melawis","Nunura","Petata","Ramin Telur","Soloman Islands Ramin"], woodDb: "ramin" },
  { name: "Redwood, Old-Growth", setting: "0.40", scientific: "Sequoia sempervirens", synonyms: ["California redwood","Coast redwood","Redwood Old-Growth","Redwood Young-Growth","Sequoia","Sequoia Pine","Vervona"], woodDb: "coast-redwood" },
  { name: "Redwood, Young-Growth", setting: "0.35", scientific: "Sequoia sempervirens", synonyms: ["California redwood","Coast redwood","Redwood Old-Growth","Redwood Young-Growth","Sequoia","Sequoia Pine","Vervona"], woodDb: "coast-redwood" },
  { name: "Roble (Tabebuia)", setting: "0.55", scientific: "Tabebuia spp." },
  { name: "Rosewood, Brazilian", setting: "0.84", scientific: "Dalbergia nigra", synonyms: ["Jacaranda de Bahia"], woodDb: "brazilian-rosewood" },
  { name: "Rosewood, Indian", setting: "0.79", scientific: "Dalbergia latifolia", woodDb: "east-indian-rosewood" },
  { name: "Rubberwood", setting: "0.51", scientific: "Hevea brasiliensis", synonyms: ["Hevea","Para Rubbertree","Plantation Hardwood"], woodDb: "rubberwood" },
  { name: "Santos Mahogany", setting: "0.95", scientific: "Myroxylon balsamum", synonyms: ["Cabreuva", "Estoraque", "Incienso", "Quina"], woodDb: "santos-mahogany" },
  { name: "Sapele", setting: "0.60", scientific: "Entandrophragma cylindricum", woodDb: "sapele" },
  { name: "Sassafras", setting: "0.46", scientific: "Sassafras albidum", synonyms: ["White Sassafras"], woodDb: "sassafras" },
  { name: "Spanish Cedar", setting: "0.44", scientific: "Cedrela odorata", synonyms: ["Acajou Rouge","Akuyari","Atoreb","Brazilian cedar","British Guiana cedar","British Honduras cedar","Cedar","Cedre rouge","Cedrela","Cedrela wood","Cedro","Cedro chino","Cedro hembra","Cedro macho","Cedro obscuro","Cedro oloroso","Cedro red","Cedro rojo","Central American cedar","Chujte","Cigar box cedar","Cigarbox cedar","Colorado cedro","Cuban cedar","Guyana cedar","Honduras cedar","Icte","Jamaican Cedar","Kalantas","Kapere","Koperi","Kurama","Kurana","Mexican Cedar","Nicaraguan Cedar","Parank","Paranka","Red Cedar","Rojas Cedar","Rosas Cedar","South American Cedar","Tabasco Cedar","Tiocuahuitl","Trinidad Cedar","Tropical Cedar","West Indian Cedar"], woodDb: "spanish-cedar" },
  { name: "Spruce, Black", setting: "0.42", scientific: "Picea mariana", synonyms: ["Bog spruce","Canadian spruce","Eastern Spruce","Shortleaf Black Spruce","Spruce","Swamp Spruce"], woodDb: "black-spruce" },
  { name: "Spruce, Engelmann", setting: "0.35", scientific: "Picea engelmannii", synonyms: ["Arizona Spruce","Balsam","Mountain Spruce","Rocky Mountain Spruce","Silver Spruce","Spruce","Western Spruce","Western White Spruce"], woodDb: "engelmann-spruce" },
  { name: "Spruce, Northern", setting: "0.36", scientific: "Picea glauca", synonyms: ["Adirondack Spruce","Blue spruce","Brunswick spruce","Canadian spruce","Cat spruce","Eastern Canadian Spruce","Eastern Spruce","Maritime Spruce","New Brunswick Spruce","Quebec spruce","Single Spruce","Skunk Spruce","St. John Spruce","Western White Spruce","Yellow Spruce"], woodDb: "white-spruce" },
  { name: "Spruce, Red", setting: "0.40", scientific: "Picea rubens", synonyms: ["Canadian red spruce","Eastern Spruce","He Balsam","Spruce","West Virginia Spruce","Yellow Spruce"], woodDb: "red-spruce" },
  { name: "Spruce, Sitka", setting: "0.40", scientific: "Picea sitchensis", synonyms: ["Californian coast spruce","Coast spruce","Menzies Spruce","Silk Spruce","Silver Spruce","Spruce","Tideland Spruce","Western Spruce","Yellow Spruce"], woodDb: "sitka-spruce" },
  { name: "Spruce, White", setting: "0.36", scientific: "Picea glauca", synonyms: ["Adirondack Spruce","Blue spruce","Brunswick spruce","Canadian spruce","Cat spruce","Eastern Canadian Spruce","Eastern Spruce","Maritime Spruce","New Brunswick Spruce","Quebec spruce","Single Spruce","Skunk Spruce","St. John Spruce","Western White Spruce","Yellow Spruce"], woodDb: "white-spruce" },
  { name: "Sweet Chestnut", setting: "0.51", scientific: "Castanea sativa", synonyms: ["Akta Kasanj","Chestnut","Edible chestnut","European chestnut","Spanish Chestnut"], woodDb: "sweet-chestnut" },
  { name: "Sweetgum", setting: "0.52", scientific: "Liquidambar styraciflua", synonyms: ["Alligator Tree","American Red Gum","American Styrax","Balsamo blanco","Bilstead","Bilsted","Blistead","Blisted","Copalillo","Copalone","Diquidambo","Estoraque","Gum Wood","Hazel Pine","Hazel Wood","Icob","Ien-gau-o","Ingano","Ko'ma","Ko'ma'liso","Liquidambar","Liquidambo","Mola","Nijte-Pijto","Nite-biito","Ocozote","Quivambaro","Sap Gum","Satin Walnut","Skchute","Slu'to'nko","Somerio","Starleaf Gum","Storax","Sweet Gum","Xochicatscahuitl","Yaga-Bizigui","Yaga-Huille","Yellow Gum"], woodDb: "sweetgum" },
  { name: "Sycamore, American", setting: "0.49", scientific: "Platanus occidentalis", synonyms: ["American Plain","American Plane","American Planetree","American Western Plane","Butterwood","Button-ball","Buttonball tree","Buttonwood","Californian button-wood","Planetree","Plataan","Sycamore"], woodDb: "sycamore" },
  { name: "SYP (Southern Yellow Pine)", setting: "0.56", scientific: "Pinus spp.", synonyms: ["Southern Pine"] },
  { name: "Tamarack", setting: "0.53", scientific: "Larix laricina", synonyms: ["Alaska Larch","American Larch","Eastern Canadian Larch","Eastern Larch","Hackmatack","Larch"], woodDb: "tamarack" },
  { name: "Tanoak", setting: "0.64", scientific: "Notholithocarpus densiflorus", woodDb: "tanoak" },
  { name: "Tatajuba", setting: "0.72", scientific: "Bagassa guianensis", woodDb: "tatajuba" },
  { name: "Tauari (Couratari)", setting: "0.53", scientific: "Couratari spp." },
  { name: "Tawa (Beilschmiedia)", setting: "0.62", scientific: "Beilschmiedia tawa" },
  { name: "Tawa (Pometia)", setting: "0.58", scientific: "Pometia spp." },
  { name: "Teak", setting: "0.57", scientific: "Tectona grandis", woodDb: "teak" },
  { name: "Tupelo, Black", setting: "0.50", scientific: "Nyssa sylvatica", synonyms: ["Chan thip","Lau Tau","Mascalwood","Pepperidge","Resak","Sourgum","Taungsagaing","Tupelo","Tupelo Gum"], woodDb: "black-tupelo" },
  { name: "Tupelo, Water", setting: "0.50", scientific: "Nyssa aquatica", synonyms: ["Bay poplar","Bowl gum","Cotton gum","Gum Cottonwood","Hazel Pine","Hickory Poplar","Olivetree","Pawpaw Gum","Sourgum","Swamp Black Gum","Swamp Poplar","Swamp Tupelo","Swamp-Gum","Tupelo","Tupelo Gum","Water Gum"], woodDb: "water-tupelo" },
  { name: "Virola", setting: "0.45", scientific: "Virola spp." },
  { name: "Walnut, Black", setting: "0.55", scientific: "Juglans nigra", synonyms: ["American Black Walnut","American Walnut","Eastern Black Walnut","Eastern Walnut","Gun-Wood","Nogal","Nogal Blanco","Nogal silvestre","Nuez Meca","Tocte","Tropical Walnut","Walnut","Walnut Tree","Wavey Black Walnut"], woodDb: "black-walnut" },
  { name: "Wenge", setting: "0.82", scientific: "Millettia laurentii", synonyms: ["Anong","Awong","Awoung","Bokonge","Bwengu","Dikela","Kiboto","Mboto","Mibotu","Monkonge","Mukonde Mutshi","Mundambi","N'gondou","N'toka","N'toko","Nson-So","Nsou-So","Otogo","Palissandre du Congo","Pallissandre","Tshikalakala","Zai-Wenge"], woodDb: "wenge" },
  { name: "Willow, Black", setting: "0.39", scientific: "Salix nigra", synonyms: ["Dudley willow","Goodding Willow","Sauce","Saule","Sauz","Southwestern Black Willow","Swamp Willow","Weide","Western Black Willow","Wilg","Willow"], woodDb: "black-willow" },
  { name: "Yellow-Poplar", setting: "0.42", scientific: "Liriodendron tulipifera", synonyms: ["American Tulipwood","American Whitewood","Blue poplar","Canadian whitewood","Canary whitewood","Canary wood","Canoe wood","Green Cypress","Hickory Poplar","Poplar","Popple","Saddle Tree","Saddletree","Tulip Poplar","Tuliptree","Tulipwood","Virginian Poplar","White Poplar","Whitewood","Yellow-Wood"], woodDb: "yellow-poplar" },
  { name: "Yew", setting: "0.63", scientific: "Taxus baccata", synonyms: ["Common yew","European yew","Idegran","Yewtree"], woodDb: "european-yew" },
  { name: "Zebrano", setting: "0.77", scientific: "Microberlinia brazzavillensis", synonyms: ["African Zebrawood","Allen Ele","Amouk","Enuk-enug","Izingana","Zebrawood","Zingana"], woodDb: "zebrawood" },
  { name: "Ziricote", setting: "0.81", scientific: "Cordia dodecandra", synonyms: ["Siricote","Zericote"], woodDb: "ziricote" }
];

// Engineered materials
const engineeredMaterials = [
  { name: "Plywood", setting: "0.57", scientific: "Various species" },
  { name: "OSB", setting: "0.62", scientific: "Various species" },
  { name: "Permacore MDF", setting: "0.70", scientific: "Various species" },
  { name: "HDF Core", setting: "0.85", scientific: "Various species" },
  { name: "Advantech™", setting: "0.70", scientific: "Various species" }
];

// Common border styling
const tableBorder = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const cellBorders = { 
  top: tableBorder, 
  bottom: tableBorder, 
  left: tableBorder, 
  right: tableBorder 
};

// Create header cell
function createHeaderCell(text) {
  return new TableCell({
    borders: cellBorders,
    shading: { fill: "2E5C8A", type: ShadingType.CLEAR },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: text, bold: true, color: "FFFFFF", size: 18 })]
    })]
  });
}

// Create data cell
function createDataCell(text, bold = false) {
  return new TableCell({
    borders: cellBorders,
    children: [new Paragraph({
      children: [new TextRun({ text: text, bold: bold, size: 16 })]
    })]
  });
}

// Split species into pages (30 species per page for readability)
function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

// Create a table for species data with 3 columns
function createSpeciesTable(speciesChunk) {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        createHeaderCell("Common Name"),
        createHeaderCell("Scientific Name"),
        createHeaderCell("Setting")
      ]
    })
  ];

  speciesChunk.forEach(species => {
    rows.push(new TableRow({
      children: [
        createDataCell(species.name),
        createDataCell(species.scientific),
        createDataCell(species.setting, true)
      ]
    }));
  });

  return new Table({
    columnWidths: [3120, 4680, 1560],
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    rows: rows
  });
}

// Build document sections
const sections = [];

// Title page
sections.push({
  properties: {
    page: {
      margin: { top: 720, right: 720, bottom: 720, left: 720 },
      size: { orientation: PageOrientation.PORTRAIT }
    }
  },
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2880, after: 480 },
      children: [new TextRun({ 
        text: "Wagner Orion 950", 
        bold: true, 
        size: 56,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
      children: [new TextRun({ 
        text: "Moisture Meter", 
        bold: true, 
        size: 48,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440 },
      children: [new TextRun({ 
        text: "Species Settings Reference", 
        size: 36,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 240 },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: "2E5C8A" },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E5C8A" }
      },
      children: [new TextRun({ 
        text: "Quick Reference Guide", 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 960, after: 120 },
      children: [new TextRun({ 
        text: "How to Use:", 
        bold: true, 
        size: 24,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "1. Locate your wood species in the alphabetical tables", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "2. Note the specific gravity (SG) setting value", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "3. Enter this value into your Wagner 950 meter using the up/down arrows", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 480, after: 120 },
      children: [new TextRun({ 
        text: "Important Notes:", 
        bold: true, 
        size: 24,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Settings are based on specific gravity at 12% moisture content", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Values represent average density for each species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Natural variation of ±10% is normal within species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Scientific names help ensure precise wood identification", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 960 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "Wagner Meters • www.wagnermeters.com • (800) 634-9961", 
        size: 18,
        color: "666666",
        font: "Arial"
      })]
    })
  ]
});

// Split species data into chunks of 30 per page
const speciesChunks = chunkArray(speciesData, 30);

speciesChunks.forEach((chunk, index) => {
  const pageChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 360 },
      children: [new TextRun({ 
        text: `Wood Species Settings (Page ${index + 1} of ${speciesChunks.length})`, 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    createSpeciesTable(chunk)
  ];

  sections.push({
    properties: {
      page: {
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: pageChildren
  });
});

// Engineered materials page
sections.push({
  properties: {
    page: {
      margin: { top: 720, right: 720, bottom: 720, left: 720 }
    }
  },
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 360 },
      children: [new TextRun({ 
        text: "Engineered Wood Products", 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    createSpeciesTable(engineeredMaterials),
    new Paragraph({
      spacing: { before: 720, after: 120 },
      children: [new TextRun({ 
        text: "Additional Information:", 
        bold: true, 
        size: 22,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• For species not listed, visit Wagner's online database at:", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 240 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "www.wagnermeters.com/specific-gravity", 
        size: 20,
        bold: true,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• The online database contains over 7,500 wood species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 480, after: 120 },
      children: [new TextRun({ 
        text: "Calibration:", 
        bold: true, 
        size: 22,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Use the included On-Demand Calibrator regularly to maintain accuracy", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Calibrate before important measurements or if meter has been dropped", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 720 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "For technical support: support@wagnermeters.com", 
        size: 18,
        color: "666666",
        font: "Arial"
      })]
    })
  ]
});

// HTML Generation Function
function createHTML(outputPath) {
  const allSpecies = [...speciesData, ...engineeredMaterials.map(m => ({ ...m, isEngineered: true }))];

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
  <meta name="theme-color" content="#ffffff" media="(prefers-color-scheme: light)">
  <meta name="theme-color" content="#1c1c1e" media="(prefers-color-scheme: dark)">
  <title>Wagner 950 Reference</title>
  <link rel="icon" href='data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y="82" font-size="82">🪵</text></svg>'>
  <style>
    :root {
      --bg: #f2f2f6;
      --panel: #ffffff;
      --border: #e3e3e8;
      --text: #1c1c1e;
      --muted: #8e8e93;
      --accent: #0a84ff;
      --line-hover: rgba(0, 0, 0, 0.035);
      --line-active: rgba(10, 132, 255, 0.08);
      --shadow: 0 1px 2px rgba(0, 0, 0, 0.06);
      --badge-bg: rgba(255, 159, 10, 0.15);
      --badge-text: #b36200;
    }

    @media (prefers-color-scheme: dark) {
      :root {
        --bg: #000000;
        --panel: #1c1c1e;
        --border: #2c2c2e;
        --text: #f2f2f7;
        --muted: #98989e;
        --line-hover: rgba(255, 255, 255, 0.05);
        --line-active: rgba(10, 132, 255, 0.16);
        --shadow: 0 1px 2px rgba(0, 0, 0, 0.4);
        --badge-bg: rgba(255, 159, 10, 0.18);
        --badge-text: #ffb340;
      }
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      min-height: 100dvh;
      display: flex;
      flex-direction: column;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Segoe UI",
        Roboto, "Helvetica Neue", sans-serif;
    }

    header {
      display: flex;
      align-items: baseline;
      flex-wrap: wrap;
      gap: 12px;
      padding: 13px 20px;
      background: var(--panel);
      border-bottom: 1px solid var(--border);
    }

    header h1 {
      margin: 0;
      font-size: 18px;
      font-weight: 700;
      letter-spacing: -0.02em;
    }

    header .tagline {
      margin: 0;
      font-size: 13px;
      color: var(--muted);
    }

    header .home {
      margin-left: auto;
      border: 1px solid var(--border);
      color: var(--muted);
      font-size: 13px;
      padding: 5px 12px;
      border-radius: 999px;
      text-decoration: none;
    }

    header .home:hover { color: var(--text); }

    main {
      flex: 1;
      width: min(1200px, 100% - 48px);
      margin: 0 auto;
      padding: 28px 0 48px;
      display: grid;
      grid-template-columns: minmax(0, 1fr) 300px;
      gap: 0 20px;
      align-items: start;
    }

    .search-box, .info-section, .table-card { grid-column: 1; }
    .search-box { grid-row: 1; }
    .info-section { grid-row: 2; }
    .table-card { grid-row: 3; }

    .detail-panel {
      grid-column: 2;
      grid-row: 1 / span 3;
      position: sticky;
      top: 20px;
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 22px;
      box-shadow: var(--shadow);
      padding: 20px 22px;
    }

    .detail-empty {
      margin: 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.5;
    }

    .detail-panel h2 {
      margin: 0;
      font-size: 20px;
      font-weight: 700;
      letter-spacing: -0.01em;
    }

    .detail-sci {
      margin: 4px 0 0;
      font-size: 14px;
      font-style: italic;
      color: var(--muted);
    }

    .detail-setting {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      margin-top: 16px;
      padding: 12px 16px;
      background: var(--bg);
      border-radius: 14px;
    }

    .detail-setting .label {
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--muted);
    }

    .detail-setting .value {
      font-size: 26px;
      font-weight: 700;
      color: var(--accent);
      font-variant-numeric: tabular-nums;
    }

    .detail-syns { margin-top: 16px; }

    .detail-syns h4 {
      margin: 0 0 6px;
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--muted);
    }

    .detail-syns p {
      margin: 0;
      font-size: 13px;
      line-height: 1.6;
      color: var(--muted);
      max-height: 40vh;
      overflow-y: auto;
    }

    .detail-link {
      display: inline-block;
      margin-top: 14px;
      font-size: 13px;
      font-weight: 600;
      color: var(--accent);
      text-decoration: none;
    }

    .detail-link:hover { text-decoration: underline; }

    @keyframes detail-flash {
      0% {
        border-color: var(--accent);
        box-shadow: 0 0 0 3px rgba(10, 132, 255, 0.25), var(--shadow);
        transform: scale(1.012);
      }
      100% {
        border-color: var(--border);
        box-shadow: var(--shadow);
        transform: scale(1);
      }
    }

    @media (prefers-reduced-motion: no-preference) {
      .detail-panel.flash {
        animation: detail-flash 0.45s ease-out;
      }
    }

    .search-box {
      position: relative;
      margin-bottom: 20px;
    }

    .search-box input {
      width: 100%;
      padding: 12px 46px 12px 18px;
      font-size: 16px;
      font-family: inherit;
      color: var(--text);
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 999px;
      outline: none;
      box-shadow: var(--shadow);
    }

    .search-box input::placeholder { color: var(--muted); }

    .search-box input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(10, 132, 255, 0.15);
    }

    .search-icon {
      position: absolute;
      right: 18px;
      top: 50%;
      transform: translateY(-50%);
      font-size: 15px;
      opacity: 0.5;
      pointer-events: none;
    }

    .info-section {
      padding: 18px 22px;
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 22px;
      box-shadow: var(--shadow);
      margin-bottom: 20px;
    }

    .info-section h3 {
      margin: 0 0 8px;
      font-size: 15px;
      font-weight: 700;
      letter-spacing: -0.01em;
    }

    .info-section ul {
      list-style: none;
      margin: 0;
      padding: 0;
    }

    .info-section li {
      padding: 4px 0 4px 24px;
      position: relative;
      font-size: 14px;
      line-height: 1.5;
      color: var(--muted);
    }

    .info-section li:before {
      content: "✓";
      position: absolute;
      left: 2px;
      color: var(--accent);
      font-weight: 600;
    }

    .table-card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 22px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .stats {
      padding: 13px 22px;
      font-size: 13px;
      color: var(--muted);
      border-bottom: 1px solid var(--border);
    }

    .stats strong {
      color: var(--text);
      font-weight: 600;
    }

    .table-container {
      max-height: 700px;
      overflow-y: auto;
    }

    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
    }

    thead {
      position: sticky;
      top: 0;
      z-index: 10;
    }

    th {
      background: var(--panel);
      color: var(--muted);
      padding: 10px 22px;
      text-align: left;
      font-weight: 600;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      border-bottom: 1px solid var(--border);
    }

    td {
      padding: 11px 22px;
      border-bottom: 1px solid var(--border);
      font-size: 14px;
    }

    tbody td:nth-child(2) {
      color: var(--muted);
      font-style: italic;
    }

    tbody tr:last-child td { border-bottom: none; }

    tbody tr { cursor: pointer; }

    tbody tr:hover {
      background-color: var(--line-hover);
    }

    tbody tr.selected,
    tbody tr.selected:hover {
      background-color: var(--line-active);
    }

    .setting-value {
      font-weight: 600;
      color: var(--accent);
      font-variant-numeric: tabular-nums;
    }

    .engineered-badge {
      display: inline-block;
      background: var(--badge-bg);
      color: var(--badge-text);
      padding: 2px 9px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.02em;
      margin-left: 8px;
      vertical-align: middle;
    }

    .no-results {
      text-align: center;
      padding: 48px 20px;
      color: var(--muted);
      font-size: 15px;
      display: none;
    }

    .no-results.show {
      display: block;
    }

    footer {
      width: min(1200px, 100% - 48px);
      margin: 0 auto;
      padding: 18px 0 calc(18px + env(safe-area-inset-bottom, 0px));
      border-top: 1px solid var(--border);
      display: flex;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 6px 14px;
      font-size: 13px;
      color: var(--muted);
    }

    footer a {
      color: var(--muted);
      text-decoration: none;
      border-bottom: 1px dotted currentColor;
    }

    footer a:hover { color: var(--text); }

    @media (max-width: 1000px) {
      main { display: block; }

      .detail-panel {
        position: static;
        margin-bottom: 20px;
      }

      .detail-panel.empty { display: none; }
    }

    @media (max-width: 768px) {
      main, footer {
        width: calc(100% - 32px);
      }

      th, td {
        padding-left: 14px;
        padding-right: 14px;
      }

      th { font-size: 11px; }

      td { font-size: 13px; }

      .stats {
        padding: 12px 14px;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>Wagner 950 Reference</h1>
    <p class="tagline">species settings for the Orion 950 moisture meter</p>
    <a class="home" href="/">More tools</a>
  </header>

  <main>
    <div class="search-box">
      <input
        type="text"
        id="searchInput"
        placeholder="Search by common name, scientific name, or setting value..."
        autocomplete="off"
      >
      <span class="search-icon">🔍</span>
    </div>

    <div class="info-section">
      <h3>How to use</h3>
      <ul>
        <li>Use the search box above to filter species by name or setting value — alternate and regional trade names match too (searching "jatoba" finds Brazilian Cherry)</li>
        <li>Locate your wood species in the table below</li>
        <li>Note the specific gravity (SG) setting value</li>
        <li>Enter this value into your Wagner 950 meter using the up/down arrows</li>
      </ul>
    </div>

    <aside class="detail-panel empty" id="detailPanel">
      <p class="detail-empty">Click a species in the table to see its details and alternate names.</p>
      <div class="detail-body" hidden>
        <h2 id="detailName"></h2>
        <p class="detail-sci" id="detailSci"></p>
        <div class="detail-setting">
          <span class="label">SG setting</span>
          <span class="value" id="detailSetting"></span>
        </div>
        <div class="detail-syns" id="detailSynsWrap">
          <h4>Also known as</h4>
          <p id="detailSyns"></p>
        </div>
        <a class="detail-link" id="detailLink" target="_blank" rel="noopener">View on The Wood Database ↗</a>
      </div>
    </aside>

    <div class="table-card">
      <div class="stats">
        Showing <strong id="visibleCount">${allSpecies.length}</strong> of <strong>${allSpecies.length}</strong> species
      </div>

      <div class="table-container">
      <table id="speciesTable">
        <thead>
          <tr>
            <th>Common Name</th>
            <th>Scientific Name</th>
            <th>Setting</th>
          </tr>
        </thead>
        <tbody id="speciesTableBody">
          ${allSpecies.map(species => `
          <tr data-name="${species.name.toLowerCase()}" data-scientific="${species.scientific.toLowerCase()}" data-setting="${species.setting}" data-synonyms="${(species.synonyms || []).join(' / ')}" data-wooddb="${species.woodDb || ''}">
            <td>
              ${species.name}${species.isEngineered ? '<span class="engineered-badge">ENGINEERED</span>' : ''}
            </td>
            <td>${species.scientific}</td>
            <td class="setting-value">${species.setting}</td>
          </tr>
          `).join('')}
        </tbody>
      </table>

      <div class="no-results" id="noResults">
        <p>No species found matching your search.</p>
        <p>Try a different search term.</p>
      </div>
      </div>
    </div>
  </main>

  <footer>
    <span>Settings are based on specific gravity at 12% moisture content · natural variation of ±10% is normal</span>
    <a href="https://www.wagnermeters.com/specific-gravity" target="_blank">wagnermeters.com/specific-gravity</a>
  </footer>

  <script>
    const searchInput = document.getElementById('searchInput');
    const tableBody = document.getElementById('speciesTableBody');
    const noResults = document.getElementById('noResults');
    const visibleCount = document.getElementById('visibleCount');
    const rows = tableBody.getElementsByTagName('tr');
    const totalCount = ${allSpecies.length};

    searchInput.addEventListener('input', function() {
      const searchTerm = this.value.toLowerCase().trim();
      let visibleRows = 0;
      let onlyVisible = null;

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const name = row.getAttribute('data-name');
        const scientific = row.getAttribute('data-scientific');
        const setting = row.getAttribute('data-setting');
        const synonyms = (row.getAttribute('data-synonyms') || '').toLowerCase();

        if (searchTerm === '' ||
            name.includes(searchTerm) ||
            scientific.includes(searchTerm) ||
            setting.includes(searchTerm) ||
            synonyms.includes(searchTerm)) {
          row.style.display = '';
          visibleRows++;
          onlyVisible = row;
        } else {
          row.style.display = 'none';
        }
      }

      // A filter narrowed to a single species: show it in the detail panel.
      // Auto-selections clear when the filter broadens again; manual ones
      // persist until their row is filtered out.
      if (visibleRows === 1 && searchTerm !== '') {
        if (selectedRow !== onlyVisible) selectRow(onlyVisible, true);
      } else if (selectedRow && (autoSelected || selectedRow.style.display === 'none')) {
        clearSelection();
      }

      visibleCount.textContent = visibleRows;

      if (visibleRows === 0) {
        noResults.classList.add('show');
        document.getElementById('speciesTable').style.display = 'none';
      } else {
        noResults.classList.remove('show');
        document.getElementById('speciesTable').style.display = 'table';
      }
    });

    // Focus search on page load
    window.addEventListener('load', function() {
      searchInput.focus();
    });

    // Clear search with Escape key
    searchInput.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        this.value = '';
        this.dispatchEvent(new Event('input'));
      }
    });

    // Detail side panel
    const detailPanel = document.getElementById('detailPanel');
    const detailBody = detailPanel.querySelector('.detail-body');
    const detailEmpty = detailPanel.querySelector('.detail-empty');
    let selectedRow = null;
    let autoSelected = false;

    tableBody.addEventListener('click', function(ev) {
      const row = ev.target.closest('tr');
      if (!row) return;
      if (selectedRow === row) {
        clearSelection();
        return;
      }
      selectRow(row, false);
    });

    function selectRow(row, isAuto) {
      if (selectedRow) selectedRow.classList.remove('selected');
      selectedRow = row;
      autoSelected = isAuto;
      row.classList.add('selected');
      showDetail(row);
    }

    function clearSelection() {
      if (selectedRow) selectedRow.classList.remove('selected');
      selectedRow = null;
      autoSelected = false;
      detailPanel.classList.add('empty');
      detailBody.hidden = true;
      detailEmpty.hidden = false;
    }

    function showDetail(row) {
      const isEngineered = !!row.querySelector('.engineered-badge');
      const name = row.cells[0].textContent.replace(/ENGINEERED\\s*$/, '').trim();
      document.getElementById('detailName').innerHTML =
        name + (isEngineered ? ' <span class="engineered-badge">ENGINEERED</span>' : '');
      document.getElementById('detailSci').textContent = row.cells[1].textContent.trim();
      document.getElementById('detailSetting').textContent = row.cells[2].textContent.trim();
      const syns = (row.getAttribute('data-synonyms') || '').split(' / ').filter(Boolean);
      document.getElementById('detailSynsWrap').hidden = syns.length === 0;
      document.getElementById('detailSyns').textContent = syns.join(', ');
      const wd = row.getAttribute('data-wooddb');
      const link = document.getElementById('detailLink');
      link.hidden = !wd;
      if (wd) link.href = 'https://www.wood-database.com/' + wd + '/';
      detailPanel.classList.remove('empty');
      detailEmpty.hidden = true;
      detailBody.hidden = false;
      detailPanel.classList.remove('flash');
      void detailPanel.offsetWidth;
      detailPanel.classList.add('flash');
    }
  </script>
</body>
</html>`;

  fs.writeFileSync(outputPath, html, 'utf8');
  return Promise.resolve();
}

// PDF Generation Functions
function createPDF(outputPath) {
  const doc = new PDFDocument({
    size: 'LETTER',
    margins: { top: 50, bottom: 50, left: 50, right: 50 }
  });

  const stream = fs.createWriteStream(outputPath);
  doc.pipe(stream);

  // Title Page
  doc.fontSize(28).fillColor('#000000').font('Helvetica-Bold')
     .text('Wagner Orion 950', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(24).text('Moisture Meter', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(18).font('Helvetica')
     .text('Species Settings Reference', { align: 'center' })
     .moveDown(2);

  // Quick Reference Guide section
  doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('Quick Reference Guide', { align: 'center' })
     .moveDown(1);

  doc.fontSize(12).fillColor('#000000').font('Helvetica-Bold')
     .text('How to Use:', { align: 'left' })
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('1. Locate your wood species in the alphabetical tables')
     .text('2. Note the specific gravity (SG) setting value')
     .text('3. Enter this value into your Wagner 950 meter using the up/down arrows')
     .moveDown(1);

  doc.fontSize(12).font('Helvetica-Bold')
     .text('Important Notes:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• Settings are based on specific gravity at 12% moisture content')
     .text('• Values represent average density for each species')
     .text('• Natural variation of ±10% is normal within species')
     .text('• Scientific names help ensure precise wood identification')
     .moveDown(2);

  doc.fontSize(9).fillColor('#666666')
     .text('Wagner Meters • www.wagnermeters.com • (800) 634-9961', { align: 'center' });

  // Species Tables
  const speciesChunks = chunkArray(speciesData, 30);

  speciesChunks.forEach((chunk, index) => {
    doc.addPage();

    doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
       .text(`Wood Species Settings (Page ${index + 1} of ${speciesChunks.length})`, { align: 'center' })
       .moveDown(0.5);

    drawSpeciesTable(doc, chunk);
  });

  // Engineered Materials Page
  doc.addPage();

  doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('Engineered Wood Products', { align: 'center' })
     .moveDown(0.5);

  drawSpeciesTable(doc, engineeredMaterials);

  doc.moveDown(2);

  doc.fontSize(11).fillColor('#000000').font('Helvetica-Bold')
     .text('Additional Information:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• For species not listed, visit Wagner\'s online database at:')
     .moveDown(0.3);

  doc.fontSize(10).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('www.wagnermeters.com/specific-gravity', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(10).fillColor('#000000').font('Helvetica')
     .text('• The online database contains over 7,500 wood species')
     .moveDown(1);

  doc.fontSize(11).font('Helvetica-Bold')
     .text('Calibration:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• Use the included On-Demand Calibrator regularly to maintain accuracy')
     .text('• Calibrate before important measurements or if meter has been dropped')
     .moveDown(2);

  doc.fontSize(9).fillColor('#666666')
     .text('For technical support: support@wagnermeters.com', { align: 'center' });

  doc.end();

  return new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });
}

function drawSpeciesTable(doc, speciesArray) {
  const tableTop = doc.y;
  const colWidths = [180, 270, 90]; // Common Name, Scientific Name, Setting
  const rowHeight = 20;
  const headerColor = '#2E5C8A';

  // Draw header row
  let y = tableTop;
  const headers = ['Common Name', 'Scientific Name', 'Setting'];

  // Header background
  doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
     .fillAndStroke(headerColor, '#000000');

  // Header text
  doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(9);
  doc.text(headers[0], 55, y + 6, { width: colWidths[0] - 10 });
  doc.text(headers[1], 55 + colWidths[0], y + 6, { width: colWidths[1] - 10 });
  doc.text(headers[2], 55 + colWidths[0] + colWidths[1], y + 6, { width: colWidths[2] - 10 });

  y += rowHeight;

  // Draw data rows
  doc.fillColor('#000000').font('Helvetica').fontSize(8);

  speciesArray.forEach((species, idx) => {
    // Alternate row backgrounds for readability
    if (idx % 2 === 0) {
      doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
         .fillAndStroke('#F5F5F5', '#000000');
    } else {
      doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
         .stroke('#000000');
    }

    // Draw cell borders
    doc.rect(50, y, colWidths[0], rowHeight).stroke('#000000');
    doc.rect(50 + colWidths[0], y, colWidths[1], rowHeight).stroke('#000000');
    doc.rect(50 + colWidths[0] + colWidths[1], y, colWidths[2], rowHeight).stroke('#000000');

    // Draw text
    doc.fillColor('#000000').font('Helvetica').fontSize(8);
    doc.text(species.name, 55, y + 6, { width: colWidths[0] - 10, lineBreak: false });
    doc.text(species.scientific, 55 + colWidths[0], y + 6, { width: colWidths[1] - 10, lineBreak: false });
    doc.font('Helvetica-Bold');
    doc.text(species.setting, 55 + colWidths[0] + colWidths[1], y + 6, { width: colWidths[2] - 10, lineBreak: false });

    y += rowHeight;
  });

  doc.y = y + 10;
}

// Create DOCX document
const docxDoc = new Document({ sections });

// Get command line arguments
const args = process.argv.slice(2);
const format = args[0] || 'all'; // 'docx', 'pdf', 'html', or 'all'

// Save to file(s)
async function generateDocuments() {
  try {
    const validFormats = ['docx', 'pdf', 'html', 'all'];

    if (!validFormats.includes(format)) {
      console.log("Usage: node wagner_950_complete.js [docx|pdf|html|all]");
      console.log("  docx - Generate only Word document");
      console.log("  pdf  - Generate only PDF document");
      console.log("  html - Generate only HTML page with search");
      console.log("  all  - Generate all formats (default)");
      return;
    }

    if (format === 'docx' || format === 'all') {
      const buffer = await Packer.toBuffer(docxDoc);
      fs.writeFileSync("Wagner_950_Reference_Sheets.docx", buffer);
      console.log("✓ DOCX file created: Wagner_950_Reference_Sheets.docx");
    }

    if (format === 'pdf' || format === 'all') {
      await createPDF("Wagner_950_Reference_Sheets.pdf");
      console.log("✓ PDF file created: Wagner_950_Reference_Sheets.pdf");
    }

    if (format === 'html' || format === 'all') {
      await createHTML("Wagner_950_Reference_Sheets.html");
      console.log("✓ HTML file created: Wagner_950_Reference_Sheets.html");
    }

    console.log("\nReference sheets created successfully with scientific names!");
  } catch (error) {
    console.error("Error generating documents:", error);
  }
}

generateDocuments();
