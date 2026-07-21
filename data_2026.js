// ============================================================
// data_2026.js — FCST 2026 extraído de "Datos HTML" (V4.xlsx)
// Generado automáticamente. NO editar manualmente.
// ============================================================

// Valor mensual total por retail (MXN)
const valorFull = {
  'AT&T': [7601182,5031587,4172841,4289142,4051383,4319526,6350496,6551785,6528680,7582562,8733235,8731866],
  'MacStore': [3427127,2193700,1406861,1516136,1368061,1528988,1529528,2616023,3422183,3870140,4240131,3725012],
  'iShop': [1685237,2195530,2869134,1371995,1834125,1150356,1358793,2050543,3951682,2317373,3152450,2215143],
  'Liverpool': [0,0,0,0,0,0,0,278981,379890,421116,469263,524848],
  'MercadoLibre': [0,0,0,0,0,0,0,0,0,0,0,0],
  'Coppel': [0,0,0,445377,0,0,0,0,0,0,0,0],
};

// Detalle por categoría/retail — Jun a Dic (pcs full, val Jun=0)
const catData = {
  'AT&T':{
    'Cables':{pcs:[9486,9978,10538,11161,12318,13757,13754],val:[0,1645422,1735936,1841030,2030677,2266112,2265415]},
    'Cases iPhone':{pcs:[898,3541,3557,3566,4055,4061,4066],val:[0,1829498,1836110,1597238,1793155,1671551,1669148]},
    'Liquid':{pcs:[813,3500,3500,3500,5000,7500,7500],val:[0,1084405,1084405,1084405,1549150,2323725,2323725]},
    'Micas':{pcs:[22,24,20,25,20,18,20],val:[0,8056,6920,7953,6843,6275,6843]},
    'Others':{pcs:[154,220,235,252,274,308,311],val:[0,51636,55274,61704,64951,73236,74399]},
    'Wall Charger':{pcs:[7972,6931,7338,7750,8558,9577,9577],val:[0,1731479,1833140,1936350,2137786,2392336,2392336]},
  },
  'MacStore':{
    'Cases AirTag':{pcs:[41,36,33,26,30,36,39],val:[0,10875,9969,7854,9062,10875,11781]},
    'Cases Airpods':{pcs:[189,300,866,830,850,893,925],val:[0,105756,219162,209031,214576,233147,245458]},
    'Cases Apple Watch':{pcs:[8,6,212,225,216,216,218],val:[0,2204,34124,36216,34768,34768,35090]},
    'Cases Macbook':{pcs:[303,208,430,302,328,392,429],val:[0,116691,238756,164965,178979,213664,233831]},
    'Cases Sleeves':{pcs:[6,4,20,12,11,13,18],val:[0,2015,10076,6046,5542,6550,9069]},
    'Cases iPad':{pcs:[460,397,608,465,490,589,642],val:[0,212837,325520,254404,264983,318547,347378]},
    'Cases iPhone':{pcs:[1072,920,1683,2688,2814,3299,3063],val:[0,478577,793109,1234128,1294633,1517978,1408909]},
    'Micas':{pcs:[302,170,340,1427,1975,1927,849],val:[0,25595,51190,519762,863754,835127,334720]},
    'Others':{pcs:[91,70,76,77,84,99,109],val:[0,17513,19654,19695,21392,25242,27718]},
    'Speakers':{pcs:[7,8,10,5,9,9,9],val:[0,10133,14838,6515,12304,12304,12304]},
    'Straps':{pcs:[429,705,1331,1437,1449,1517,1541],val:[0,547332,888651,959648,959957,1021739,1049347]},
    'Wall Charger':{pcs:[0,0,14,5,13,13,12],val:[0,0,10974,3919,10190,10190,9407]},
  },
  'iShop':{
    'Cases AirTag':{pcs:[37,59,80,36,84,100,51],val:[0,13710,18590,8365,19519,23237,11851]},
    'Cases Airpods':{pcs:[165,104,173,190,144,90,90],val:[0,31903,57368,60329,44487,25889,25889]},
    'Cases Macbook':{pcs:[417,378,1400,777,797,1482,1077],val:[0,189297,647993,353434,376595,688906,489942]},
    'Cases Sleeves':{pcs:[8,17,21,12,33,12,12],val:[0,7320,9043,5167,14210,5167,5167]},
    'Cases iPad':{pcs:[801,690,1775,806,751,1556,821],val:[0,322762,816732,376721,352607,718551,400508]},
    'Cases iPhone':{pcs:[681,1375,650,6699,2644,2844,2189],val:[0,570052,259994,2933744,1107024,1168792,939326]},
    'Others':{pcs:[0,30,27,38,10,40,38],val:[0,6971,6274,8830,2324,9295,8830]},
    'Straps':{pcs:[399,378,448,389,696,942,582],val:[0,216778,234549,205092,400607,512613,333630]},
  },
  'Liverpool':{
    'Cases Airpods':{pcs:[0,0,300,354,418,494,582],val:[0,0,108464,127988,151128,178606,210422]},
    'Cases iPhone':{pcs:[0,0,140,300,300,300,300],val:[0,0,72340,138224,138224,138224,138224]},
    'Straps':{pcs:[0,0,190,220,255,295,341],val:[0,0,98177,113678,131764,152433,176202]},
  },
};

// Budget/FCST 2026 — Valor mensual por retail (MXN)
const avb2026Val = {
  'AT&T': [7601182,5031587,4172841,4289142,4051383,4319526,6350496,6551785,6528680,7582562,8733235,8731866],
  'MacStore': [3427127,2193700,1406861,1516136,1368061,1528988,1529528,2616023,3422183,3870140,4240131,3725012],
  'iShop': [1685237,2195530,2869134,1371995,1834125,1150356,1358793,2050543,3951682,2317373,3152450,2215143],
  'Liverpool': [0,0,0,0,0,0,0,278981,379890,421116,469263,524848],
  'MercadoLibre': [0,0,0,0,0,0,0,0,0,0,0,0],
  'Coppel': [0,0,0,445377,0,0,0,0,0,0,0,0],
};

// Budget/FCST 2026 — Piezas mensual por retail
const avb2026Pcs = {
  'AT&T': [30710,20985,18176,19397,18409,19345,24194,25188,26254,30225,35221,35228],
  'MacStore': [7961,4633,3011,2947,2520,2908,2824,5623,7499,8269,9003,7854],
  'iShop': [3969,4928,6123,3132,4095,2508,3031,4574,8947,5159,7066,4860],
  'Liverpool': [0,0,0,0,0,0,0,630,874,973,1089,1223],
  'MercadoLibre': [0,0,0,0,0,0,0,0,0,0,0,0],
  'Coppel': [0,0,0,1470,0,0,0,0,0,0,0,0],
};

// Valor trimestral 2026 por retail (Q1=Ene-Mar … Q4=Oct-Dic)
const qValorAll = {
  'AT&T': [16805610,12660051,19430961,25047663],
  'MacStore': [7027688,4413185,7567734,11835283],
  'iShop': [6749901,4356476,7361018,7684966],
  'Liverpool': [0,0,658871,1415227],
  'MercadoLibre': [0,0,0,0],
  'Coppel': [0,445377,0,0],
};

// Piezas trimestrales 2026 por retail
const qPiezas = {
  'AT&T': [69871,57151,75636,100674],
  'MacStore': [15605,8375,15946,25126],
  'iShop': [15020,9735,16552,17085],
  'Liverpool': [0,0,1504,3285],
  'MercadoLibre': [0,0,0,0],
  'Coppel': [0,1470,0,0],
};

// Detalle completo por marca → retail → categoría → {pcs, val} mensual
const marcaRaw = {
  'Nomad':{
    'iShop':{
      'Cases Airpods':{pcs:{Jan:447,Feb:240,Mar:319,Apr:113,May:275,Jun:157,Jul:81,Aug:160,Sep:160,Oct:114,Nov:60,Dec:60},val:{Jan:153947,Feb:82656,Mar:109864,Apr:38917,May:94710,Jun:54071,Jul:27897,Aug:55104,Sep:55104,Oct:39262,Nov:20664,Dec:20664}},
      'Cases iPhone':{pcs:{Jan:388,Feb:343,Mar:452,Apr:145,May:344,Jun:154,Jul:246,Aug:228,Sep:1170,Oct:670,Nov:710,Dec:570},val:{Jan:202179,Feb:178730,Mar:235529,Apr:75557,May:179251,Jun:80247,Jul:128186,Aug:118805,Sep:586565,Oct:330314,Nov:349263,Dec:282944}},
      'Straps':{pcs:{Jan:331,Feb:684,Mar:859,Apr:485,May:711,Jun:399,Jul:378,Aug:448,Sep:389,Oct:696,Nov:942,Dec:582},val:{Jan:188780,Feb:431732,Mar:479934,Apr:278024,May:413144,Jun:222803,Jul:216778,Aug:234549,Sep:205092,Oct:400607,Nov:512613,Dec:333630}},
    },
    'Liverpool':{
      'Cases Airpods':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:300,Sep:354,Oct:418,Nov:494,Dec:582},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:108464,Sep:127988,Oct:151128,Nov:178606,Dec:210422}},
      'Cases iPhone':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:140,Sep:300,Oct:300,Nov:300,Dec:300},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:72340,Sep:138224,Oct:138224,Nov:138224,Dec:138224}},
      'Straps':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:190,Sep:220,Oct:255,Nov:295,Dec:341},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:98177,Sep:113678,Oct:131764,Nov:152433,Dec:176202}},
    },
    'ATT':{
      'Cases iPhone':{pcs:{Jan:4,Feb:1,Mar:1,Apr:3,May:1,Jun:1,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:2067,Feb:517,Mar:517,Apr:1551,May:517,Jun:517,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
    'MacStore':{
      'Cables':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Cases AirTag':{pcs:{Jan:1,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:482,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Cases Airpods':{pcs:{Jan:57,Feb:59,Mar:168,Apr:183,May:156,Jun:170,Jul:159,Aug:182,Sep:170,Oct:176,Nov:211,Dec:232},val:{Jan:20905,Feb:24163,Mar:79916,Apr:88114,May:74976,Jun:81846,Jul:76423,Aug:87754,Sep:81967,Oct:84861,Nov:101736,Dec:111861}},
      'Cases iPhone':{pcs:{Jan:508,Feb:283,Mar:199,Apr:232,May:234,Jun:277,Jul:236,Aug:255,Sep:260,Oct:313,Nov:313,Dec:313},val:{Jan:306244,Feb:170605,Mar:119967,Apr:139860,May:141065,Jun:166986,Jul:142270,Aug:153724,Sep:121926,Oct:146061,Nov:146061,Dec:146061}},
      'Straps':{pcs:{Jan:697,Feb:559,Mar:416,Apr:405,May:380,Jun:429,Jul:384,Aug:462,Sep:475,Oct:491,Nov:543,Dec:553},val:{Jan:552350,Feb:464077,Mar:378667,Apr:367246,May:359927,Jun:370982,Jul:365617,Aug:430610,Sep:435996,Oct:446705,Nov:501853,Dec:523839}},
      'Wall Charger':{pcs:{Jan:12,Feb:34,Mar:16,Apr:6,May:4,Jun:0,Jul:0,Aug:14,Sep:5,Oct:13,Nov:13,Dec:12},val:{Jan:7596,Feb:22187,Mar:11818,Apr:3980,May:2412,Jun:0,Jul:0,Aug:10974,Sep:3919,Oct:10190,Nov:10190,Dec:9407}},
    },
  },
  'Pitaka':{
    'MacStore':{
      'Cases iPhone':{pcs:{Jan:339,Feb:201,Mar:96,Apr:188,May:150,Jun:177,Jul:114,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:206053,Feb:122137,Mar:58355,Apr:114300,May:91391,Jun:107185,Jul:71138,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Others':{pcs:{Jan:4,Feb:8,Mar:0,Apr:0,May:0,Jun:1,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:3014,Feb:3857,Mar:0,Apr:0,May:0,Jun:482,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Straps':{pcs:{Jan:8,Feb:4,Mar:1,Apr:0,May:0,Jun:0,Jul:2,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:5064,Feb:2533,Mar:603,Apr:0,May:0,Jun:0,Jul:1206,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
  },
  'Quikcell':{
    'ATT':{
      'Cables':{pcs:{Jan:11224,Feb:9129,Mar:8501,Apr:9707,May:9383,Jun:9486,Jul:9978,Aug:10538,Sep:11161,Oct:12318,Nov:13757,Dec:13754},val:{Jan:1836215,Feb:1577163,Mar:1426273,Apr:1588019,May:1524565,Jun:1558080,Jul:1645422,Aug:1735936,Sep:1841030,Oct:2030677,Nov:2266112,Dec:2265415}},
      'Cases iPhone':{pcs:{Jan:45,Feb:74,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:9287,Feb:15273,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Liquid':{pcs:{Jan:1834,Feb:1252,Mar:830,Apr:991,May:778,Jun:813,Jul:3500,Aug:3500,Sep:3500,Oct:5000,Nov:7500,Dec:7500},val:{Jan:568228,Feb:387907,Mar:257159,Apr:307042,May:241048,Jun:251892,Jul:1084405,Aug:1084405,Sep:1084405,Oct:1549150,Nov:2323725,Dec:2323725}},
      'Micas':{pcs:{Jan:44,Feb:14,Mar:23,Apr:22,May:10,Jun:14,Jul:8,Aug:4,Sep:14,Oct:5,Nov:3,Dec:5},val:{Jan:12495,Feb:3976,Mar:6532,Apr:6247,May:2840,Jun:3976,Jul:2272,Aug:1136,Sep:3976,Oct:1420,Nov:852,Dec:1420}},
      'Others':{pcs:{Jan:319,Feb:275,Mar:192,Apr:161,May:161,Jun:136,Jul:213,Aug:226,Sep:238,Oct:263,Nov:295,Dec:295},val:{Jan:74085,Feb:63866,Mar:44590,Apr:37391,May:37391,Jun:31585,Jul:49467,Aug:52486,Sep:55273,Oct:61079,Nov:68511,Dec:68511}},
      'Wall Charger':{pcs:{Jan:14078,Feb:8677,Mar:7613,Apr:7656,May:7211,Jun:7972,Jul:6931,Aug:7338,Sep:7750,Oct:8558,Nov:9577,Dec:9577},val:{Jan:3474868,Feb:2187718,Mar:1919327,Apr:1912942,May:1801369,Jun:2004614,Jul:1731479,Aug:1833140,Sep:1936350,Oct:2137786,Nov:2392336,Dec:2392336}},
    },
    'MacStore':{
      'Cables':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Cases iPhone':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Liquid':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Micas':{pcs:{Jan:2016,Feb:788,Mar:723,Apr:387,May:238,Jun:302,Jul:170,Aug:340,Sep:444,Oct:149,Nov:170,Dec:182},val:{Jan:305352,Feb:120632,Mar:109808,Apr:59065,May:36639,Jun:45628,Jul:25595,Aug:51190,Sep:66849,Oct:22433,Nov:25596,Dec:27403}},
      'Others':{pcs:{Jan:62,Feb:78,Mar:54,Apr:57,May:50,Jun:49,Jul:31,Aug:43,Sep:46,Oct:49,Nov:58,Dec:64},val:{Jan:13199,Feb:16605,Mar:11496,Apr:12135,May:10644,Jun:10432,Jul:6600,Aug:9154,Sep:9793,Oct:10432,Nov:12348,Dec:13625}},
      'Wall Charger':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
  },
  'Raptic':{
    'MacStore':{
      'Cases Apple Watch':{pcs:{Jan:124,Feb:54,Mar:20,Apr:6,May:7,Jun:8,Jul:6,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:49802,Feb:17602,Mar:6603,Apr:1565,May:1928,Jun:3471,Jul:2204,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
  },
  'Speck':{
    'Coppel':{
      'Cases iPhone':{pcs:{Jan:0,Feb:0,Mar:0,Apr:1470,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:445377,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
    'iShop':{
      'Cases AirTag':{pcs:{Jan:36,Feb:51,Mar:128,Apr:36,May:85,Jun:37,Jul:59,Aug:80,Sep:36,Oct:84,Nov:100,Dec:51},val:{Jan:8365,Feb:11851,Mar:29743,Apr:8365,May:19751,Jun:8598,Jul:13710,Aug:18590,Sep:8365,Oct:19519,Nov:23237,Dec:11851}},
      'Cases Macbook':{pcs:{Jan:229,Feb:561,Mar:800,Apr:267,May:565,Jun:254,Jul:260,Aug:1200,Sep:707,Oct:677,Nov:1282,Dec:1037},val:{Jan:103457,Feb:258377,Mar:370234,Apr:125963,May:264516,Jun:117346,Jul:120684,Aug:531699,Sep:312731,Oct:306819,Nov:572612,Dec:466683}},
      'Cases Sleeves':{pcs:{Jan:0,Feb:0,Mar:35,Apr:9,May:0,Jun:8,Jul:17,Aug:21,Sep:12,Oct:33,Nov:12,Dec:12},val:{Jan:0,Feb:0,Mar:15071,Apr:3875,May:0,Jun:3445,Jul:7320,Aug:9043,Sep:5167,Oct:14210,Nov:5167,Dec:5167}},
      'Cases iPad':{pcs:{Jan:1038,Feb:1133,Mar:1389,Apr:548,May:1058,Jun:659,Jul:547,Aug:1625,Sep:699,Oct:644,Nov:1411,Dec:601},val:{Jan:449595,Feb:488992,Mar:598451,Apr:237521,May:447559,Jun:285061,Jul:236530,Aug:726279,Sep:312197,Oct:288083,Nov:631113,Dec:267844}},
      'Cases iPhone':{pcs:{Jan:617,Feb:912,Mar:478,Apr:512,May:481,Jun:316,Jul:632,Aug:230,Sep:2324,Oct:899,Nov:929,Dec:744},val:{Jan:238802,Feb:353435,Mar:185245,Apr:197955,May:185631,Jun:122463,Jul:244538,Aug:78427,Sep:922914,Oct:345062,Nov:346990,Dec:284993}},
      'Others':{pcs:{Jan:31,Feb:6,Mar:35,Apr:30,May:64,Jun:0,Jul:30,Aug:27,Sep:38,Oct:10,Nov:40,Dec:38},val:{Jan:7203,Feb:1394,Mar:8133,Apr:6971,May:14872,Jun:0,Jul:6971,Aug:6274,Sep:8830,Oct:2324,Nov:9295,Dec:8830}},
      'Sleeves':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
      'Speakers':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0}},
    },
    'ATT':{
      'Cases iPhone':{pcs:{Jan:3093,Feb:1489,Mar:966,Apr:819,May:844,Jun:889,Jul:3541,Aug:3557,Sep:2479,Oct:2724,Nov:2388,Dec:2373},val:{Jan:1595319,Feb:766500,Mar:497083,Apr:421125,May:435492,Jun:456260,Jul:1829498,Aug:1836110,Sep:1157532,Oct:1256265,Nov:996856,Dec:986281}},
      'Micas':{pcs:{Jan:14,Feb:26,Mar:12,Apr:15,May:8,Jun:8,Jul:16,Aug:16,Sep:11,Oct:15,Nov:15,Dec:15},val:{Jan:5061,Feb:9400,Mar:4338,Apr:5423,May:2892,Jun:2892,Jul:5784,Aug:5784,Sep:3977,Oct:5423,Nov:5423,Dec:5423}},
      'Others':{pcs:{Jan:28,Feb:29,Mar:25,Apr:12,May:7,Jun:18,Jul:7,Aug:9,Sep:14,Oct:11,Nov:13,Dec:16},val:{Jan:9606,Feb:9449,Mar:10305,Apr:3718,May:2169,Jun:5577,Jul:2169,Aug:2788,Sep:6431,Oct:3872,Nov:4725,Dec:5888}},
    },
    'MacStore':{
      'Cases AirTag':{pcs:{Jan:48,Feb:26,Mar:22,Apr:38,May:34,Jun:41,Jul:36,Aug:33,Sep:26,Oct:30,Nov:36,Dec:39},val:{Jan:14500,Feb:7854,Mar:6646,Apr:11479,May:10271,Jun:12385,Jul:10875,Aug:9969,Sep:7854,Oct:9062,Nov:10875,Dec:11781}},
      'Cases Macbook':{pcs:{Jan:274,Feb:254,Mar:152,Apr:213,May:141,Jun:182,Jul:130,Aug:285,Sep:218,Oct:238,Nov:286,Dec:313},val:{Jan:138044,Feb:127968,Mar:76579,Apr:107312,May:71037,Jun:91693,Jul:65495,Aug:143585,Sep:109831,Oct:119907,Nov:144090,Dec:157693}},
      'Cases Sleeves':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:3,Jun:6,Jul:4,Aug:20,Sep:12,Oct:11,Nov:13,Dec:18},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:1511,Jun:3023,Jul:2015,Aug:10076,Sep:6046,Oct:5542,Nov:6550,Dec:9069}},
      'Cases iPad':{pcs:{Jan:1538,Feb:865,Mar:352,Apr:394,May:374,Jun:393,Jul:340,Aug:510,Sep:363,Oct:397,Nov:477,Dec:519},val:{Jan:768759,Feb:435595,Mar:178955,Apr:197492,May:187821,Jun:199861,Jul:172304,Aug:255833,Sep:181873,Oct:198852,Nov:238905,Dec:259914}},
      'Cases iPhone':{pcs:{Jan:1521,Feb:859,Mar:460,Apr:425,May:364,Jun:403,Jul:264,Aug:423,Sep:829,Oct:784,Nov:997,Dec:953},val:{Jan:690476,Feb:390068,Mar:207704,Apr:191392,May:162179,Jun:182363,Jul:118872,Aug:190654,Sep:376242,Oct:356800,Nov:454939,Dec:434480}},
      'Others':{pcs:{Jan:64,Feb:43,Mar:46,Apr:30,May:33,Jun:41,Jul:39,Aug:33,Sep:31,Oct:35,Nov:41,Dec:45},val:{Jan:19019,Feb:13355,Mar:13216,Apr:8702,May:10362,Jun:12341,Jul:10913,Aug:10500,Sep:9902,Oct:10960,Nov:12894,Dec:14093}},
      'Speakers':{pcs:{Jan:11,Feb:6,Mar:4,Apr:3,May:2,Jun:7,Jul:8,Aug:10,Sep:5,Oct:9,Nov:9,Dec:9},val:{Jan:13752,Feb:7600,Mar:4342,Apr:3981,May:2534,Jun:9047,Jul:10133,Aug:14838,Sep:6515,Oct:12304,Nov:12304,Dec:12304}},
    },
  },
  'Tech21':{
    'iShop':{
      'Cases Airpods':{pcs:{Jan:0,Feb:0,Mar:0,Apr:81,May:39,Jun:8,Jul:23,Aug:13,Sep:30,Oct:30,Nov:30,Dec:30},val:{Jan:0,Feb:0,Mar:0,Apr:14109,May:6793,Jun:1393,Jul:4006,Aug:2264,Sep:5225,Oct:5225,Nov:5225,Dec:5225}},
      'Cases Macbook':{pcs:{Jan:0,Feb:0,Mar:302,Apr:112,May:0,Jun:163,Jul:118,Aug:200,Sep:70,Oct:120,Nov:200,Dec:40},val:{Jan:0,Feb:0,Mar:175604,Apr:65125,May:0,Jun:94780,Jul:68613,Aug:116294,Sep:40703,Oct:69776,Nov:116294,Dec:23259}},
      'Cases iPad':{pcs:{Jan:0,Feb:0,Mar:449,Apr:139,May:74,Jun:142,Jul:143,Aug:150,Sep:107,Oct:107,Nov:145,Dec:220},val:{Jan:0,Feb:0,Mar:270756,Apr:83820,May:44623,Jun:85629,Jul:86232,Aug:90453,Sep:64524,Oct:64524,Nov:87438,Dec:132664}},
      'Cases iPhone':{pcs:{Jan:852,Feb:998,Mar:877,Apr:655,May:399,Jun:211,Jul:497,Aug:192,Sep:3205,Oct:1075,Nov:1205,Dec:875},val:{Jan:332909,Feb:388363,Mar:390570,Apr:235793,May:163275,Jun:74520,Jul:197328,Aug:62762,Sep:1424265,Oct:431648,Nov:472539,Dec:371389}},
    },
    'ATT':{
      'Cases iPhone':{pcs:{Jan:27,Feb:19,Mar:13,Apr:11,May:6,Jun:8,Jul:0,Aug:0,Sep:1087,Oct:1331,Nov:1673,Dec:1693},val:{Jan:13951,Feb:9818,Mar:6717,Apr:5684,May:3100,Jun:4133,Jul:0,Aug:0,Sep:439706,Oct:536890,Nov:674695,Dec:682867}},
    },
    'MacStore':{
      'Cases Airpods':{pcs:{Jan:71,Feb:45,Mar:26,Apr:13,May:5,Jun:19,Jul:7,Aug:24,Sep:23,Oct:21,Nov:23,Dec:23},val:{Jan:16077,Feb:10189,Mar:5887,Apr:2944,May:1132,Jun:4302,Jul:1585,Aug:5434,Sep:5208,Oct:4755,Nov:5208,Dec:5208}},
      'Cases Macbook':{pcs:{Jan:84,Feb:107,Mar:79,Apr:111,May:139,Jun:121,Jul:78,Aug:145,Sep:84,Oct:90,Nov:106,Dec:116},val:{Jan:55134,Feb:70229,Mar:51851,Apr:72855,May:91233,Jun:79420,Jul:51196,Aug:95171,Sep:55134,Oct:59072,Nov:69574,Dec:76138}},
      'Cases iPad':{pcs:{Jan:0,Feb:0,Mar:12,Apr:63,May:64,Jun:67,Jul:57,Aug:98,Sep:102,Oct:93,Nov:112,Dec:123},val:{Jan:0,Feb:0,Mar:8533,Apr:44799,May:45510,Jun:47643,Jul:40533,Aug:69687,Sep:72531,Oct:66131,Nov:79642,Dec:87464}},
      'Cases iPhone':{pcs:{Jan:522,Feb:360,Mar:165,Apr:193,May:142,Jun:215,Jul:158,Aug:208,Sep:924,Oct:627,Nov:679,Dec:647},val:{Jan:241305,Feb:166444,Mar:75915,Apr:88915,May:65489,Jun:99898,Jul:75164,Aug:95966,Sep:426400,Oct:289558,Nov:313400,Dec:298508}},
      'Micas':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:983,Oct:1826,Nov:1757,Dec:667},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:0,Sep:452913,Oct:841321,Nov:809531,Dec:307317}},
    },
  },
  'Uniq':{
    'MacStore':{
      'Cases Airpods':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:134,Aug:660,Sep:637,Oct:653,Nov:659,Dec:670},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:27748,Aug:125974,Sep:121856,Oct:124960,Nov:126203,Dec:128389}},
      'Cases Apple Watch':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:212,Sep:225,Oct:216,Nov:216,Dec:218},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:0,Aug:34124,Sep:36216,Oct:34768,Nov:34768,Dec:35090}},
      'Cases iPhone':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:148,Aug:797,Sep:675,Oct:1090,Nov:1310,Dec:1150},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:71133,Aug:352765,Sep:309560,Oct:502214,Nov:603578,Dec:529860}},
      'Straps':{pcs:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:319,Aug:869,Sep:962,Oct:958,Nov:974,Dec:988},val:{Jan:0,Feb:0,Mar:0,Apr:0,May:0,Jun:0,Jul:180509,Aug:458041,Sep:523652,Oct:513252,Nov:519886,Dec:525508}},
    },
  },
};

// TOTALES 2026: $133,460,011 MXN | 433,035 piezas
// Retails activos: ATT, MacStore, iShop, Liverpool, Coppel
// Marcas: Nomad, Pitaka, Quikcell, Raptic, Speck, Tech21, Uniq