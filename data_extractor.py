+
    Ó>Âiøš  ã                   ór  € R t ^ RIt^ RIt^ RIt^ RIt^ RIt^ RI Ht ^ RI	H
t
 RR.t
Rt. R-Ot
. R.Ot. R/Ot0 R0mtR  R ltR1R	 R
 lltR
 R ltR
 R ltR R ltR R ltR R ltR R ltR R ltR R ltR R ltR R ltR R  ltR! R" ltR2R# R$ lltR3R% R& llt R3R' R( llt!R3R) R* llt"R3R+ R, llt#R# )4uõ   
data_extractor.py
Centro Floricultor de Baja California
- Hojas WK  â†’ Excel en OneDrive (pandas + requests)
- Hojas PR  â†’ Google Sheets (gspread + service account)
- Hojas MP  â†’ Google Sheets (gspread + service account) â€” MANTENIMIENTO
N©Ú BytesIO)Ú
Credentialsz5https://www.googleapis.com/auth/spreadsheets.readonlyz.https://www.googleapis.com/auth/drive.readonlyz‚https://pacificafarms-my.sharepoint.com/:x:/g/personal/anahi_mora_cfbc_co/IQAQCb79SzHtRrTQR71pSNQcASOWqFXyeGGzEhUcT9FRRJ4?e=ClxLCNc                ó2   € V ^8„  d   QhR\         R,           /# )é   ÚreturnNr   )Úformats   "Údata_extractor (17).pyÚ__annotate__r
   <   s   € ÷ 
ñ 
œ 4ñ 
ó    c                 óø   € \         P                  RR4      p  \        P                  ! V ^R7      pVP	                  4        \
        VP                  4      #   \         d   p\        RT 24        Rp?R# Rp?ii ; i)zCDescarga el archivo .xlsx desde OneDrive y lo retorna como BytesIO.z?e=z?download=1&e=)Ú timeoutu"   âŒ Error descargando el archivo: N)	ÚONEDRIVE_URLÚ replaceÚrequestsÚgetÚraise_for_statusr   Ú contentÚ	ExceptionÚprint)Údownload_urlÚresponseÚes      r	   Údescargar_excelr   <   sl   € ô  ×'Ñ'¨Ð/?Ó@€LðÜ—<’< °bÔ9ˆØ×!Ñ!Ô#Üx×'Ñ'Ó(Ð(øÜ
ô Ü
Ð2°1°#Ð6Ô7Ýûðús   ˜<A Á
A9Á A4Á4A9c          
      ó~   € V ^8„  d   QhR\         P                  R\        R\        R\        R\        \        ,          /# )r   ÚxlsÚtituloÚ
rango_filasÚ
rango_colsr    )ÚpdÚ	ExcelFileÚstrÚintÚlist)r   s   "r	   r
   r
   I   s9   € ÷ ñ ”B—L‘Lð ¬#ð ¼Cð ÜðÜ(,¬T­
ñr
   c                 ó>  €  \         P                  ! V VRVR7      P                  R4      pVP                  ^,          V8”  d   VP                  RRV13,          pVP
                  P
                  4       #   \         d   p\        RT RT 24       . u Rp?# Rp?ii ; i) uº   
Lee una hoja del ExcelFile y la retorna como lista de listas.
Las celdas vacÃ­as / NaN se convierten a "".
rango_filas / rango_cols limitan cuÃ¡nto leer (equivalente al rango A1:AI60).
N)Ú
sheet_nameÚheaderÚnrowsÚ :NNNu      âš ï¸  Error leyendo hoja 'z': )	r   Ú
read_excelÚfillnaÚshapeÚilocÚvaluesÚtolistr   r   )r   r   r   r   Údfr   s   &&&&  r	   Ú
_leer_hojar0   I   s•   € ð
Ü
]Š]ØØØØô	
÷
 
‰&‹*ð
 	
ð 8‰8A;˜Ô
#Ø—‘˜˜K˜Z˜K˜Õ(ˆBØy‰y×ÑÓ!Ð!øÜ
ô Ü
Ð/°¨x°s¸1¸#Ð>Ô?Ø	ûðús   ‚A1A4 Á4
BÁ?BÂBÂBc                ó$   € V ^8„  d   QhR\         /# ©r   Ús©r!   )r   s   "r	   r
   r
   a   s   € ÷ ñ ”#ñ r
   c                 ó2  € \        V 4      P                  4       P                  4       p R V 9   d   R# RV 9   d   R# RV 9   g   RV 9   d   R# R V 9   d   R# R	V 9   d   R	# R
V 9   d   R
# RV 9   d   R
# RV 9   d   R# RV 9   d   R# RV 9   d   RV 9  d
   RV 9  d   R# R# )ÚPROPú Prop-RMÚPOSCOúPosCo-RMzCAMPO-VIzCAMPO-IVzCampo-VIÚALBAHACAú
Albahaca-RMÚHOOPSÚ	CHRISTINAÚ	Christinaz
CECILIA 25ú
Cecilia 25Ú CECILIAÚ CeciliaÚISABELÚ IsabelaÚCAMPOÚVIÚIVúCampo-RMN)r!   ÚupperÚstrip©r3   s   &r	   Ú
norm_ranchrK   a   s‹   € Ü
ˆA‹‰‹×ÑÓ€AØ 
„{Á	Ø !„|Á
Ø Q„˜*¨œ/Á
Ø Q„Á
Ø !„|Á Ø aÔ Á
Ø qÔ ÁØ A„~Á	Ø 1„}Á	Ø !„|˜ Aœ
¨$°a¬-Á
Ù
r
   c                ó0   € V ^8„  d   QhR\         R\         /# )r   r3   r    r4   )r   s   "r	   r
   r
   p   s   € ÷ 
ñ 
”#ð 
œ#ñ 
r
   c                 ó  € \        T ;'       g    R 4      P                  4       P                  4       p \        P                  ! RV 4      P
                  RR4      P
                  R4      p \        P                  ! RRV 4      p V # )r(   ÚNFKDÚasciiÚignoreú\s+Ú )	r!   rI   rH   Ú
unicodedataÚ	normalizeÚencodeÚdecodeÚreÚsubrJ   s   &r	   Ú
_norm_textrY   p   sd   € Ü
ˆAGˆG‹×ÑÓ×"Ñ"Ó$€AÜ×Ò˜f aÓ(×/Ñ/° ¸ÓB×IÑIÈ'ÓR€AÜ
Šˆvs˜AÓ€AØ
€Hr
   c                ó$   € V ^8„  d   QhR\         /# r2   r4   )r   s   "r	   r
   r
   w   s   € ÷ ñ ”Cñ r
   c                 ó`   € \        V 4      p R V 9   d   R# RV 9   d   R# RV 9   g   RV 9   d   R# R # )zCOSTO DE MATÚ	materialsz
COSTO DE SERVÚservicesz
COSTO DE MANOzMANO DE OBRAÚlaborN©rY   rJ   s   &r	   Únorm_sectionr`   w   s6   € Ü1‹
€AØ ˜Ô ÙØ ˜!Ô ÙØ ˜!Ô ˜~°Ô2ÙÙ
r
   c                ó$   € V ^8„  d   QhR\         /# r2   r4   )r   s   "r	   r
   r
   ‚   s   € ÷ 
ñ 
œñ 
r
   c                 ó  € \        V 4      p R V 9   d
   RV 9   d   R# V P                  R4      '       d   R# RV 9   d   R# RV 9   d   R # RV 9   d   R	# R
V 9   d   R
# RV 9   g   R
V 9   d   R# RV 9   d   R# RV 9   d   R# RV 9   d   R# RV 9   d   R# R# )ÚDESINFECCIONÚFERTILIZúDESINFECCION Y FERTILIZACIONÚ
AMPLIACIONÚ CULTIVOúCULTIVO TIERRA, CHAROLASzMATERIAL VEGúMATERIAL VEGETALÚ
PREPARACIONúPREPARACION DE SUELOÚFERTILIZANTEÚ
FERTILIZANTESÚ SANIDADÚ
PLAGUICIDAúDESINFECCION / PLAGUICIDASÚ
MANTENIMIENTOÚ	EXPANSIONúEXPANSION CECILIA 25Ú
RENOVACIONúRENOVACION DE SIEMBRAzMATERIAL DE EMPúMATERIAL DE EMPAQUEN)rY   Ú
startswithrJ   s   &r	   Únorm_material_catrx   ‚   s–   € Ü1‹
€AØ ˜Ô ˜z¨QœÑ8VØ ‡||L× !Ò !¹Ø A„~Ñ9SØ ˜Ô Ñ9KØ ˜Ô Ñ9OØ ˜Ô ¹Ø A„~˜¨Ô*Ñ8TØ ˜!Ô ¹Ø aÔ Ñ9OØ qÔ Ñ9PØ ˜AÔ Ñ9NÙ
r
   c                ó$   € V ^8„  d   QhR\         /# r2   r4   )r   s   "r	   r
   r
   ’   s   € ÷ ñ œñ r
   c                 ó*  € \        V 4      p R V 9   d   R# RV 9   d
   RV 9   d   R# RV 9   d   R# R V 9   d   R# R	V 9   d
   R
V 9   d   R
# RV 9   d   R
V 9   d
   RV 9   d   R# RV 9   d
   R
V 9   d   R# RV 9   d   RV 9   g   RV 9   d   RV 9   g   RV 9   d   R# R# )ÚELECTRICÚELECTRICIDADÚFLETEÚACARREúFLETES Y ACARREOSÚEXPORTúGASTOS DE EXPORTACIONÚFITOúCERTIFICADO FITOSANITARIOÚ
TRANSPORTEÚPERSONALúTRANSPORTE DE PERSONALÚCOMPRAÚFLORÚTERCERúCOMPRA DE FLOR A TERCEROSÚCOMIDAúCOMIDA PARA EL PERSONALÚTELÚROzR/OÚRTAÚALIMúRO, TEL, RTA.ALIMNr_   rJ   s   &r	   Únorm_service_catr’   ’   s•   € Ü1‹
€AØ Q„ÙØ !„|˜ Aœ
Ù"Ø 1„}Ù&Ø 
„{Ù*Ø qÔ ˜Z¨1œ_Ù'Ø 1„}˜ 1œ¨°Q¬Ù*Ø 1„}˜ qœÙ(Ø „zt˜q”y E¨Q¤J°U¸a´ZÀ6ÈQÄ;Ù"Ù
r
   c                ó$   € V ^8„  d   QhR\         /# r2   r4   )r   s   "r	   r
   r
   §   s   € ÷  ñ  ”ñ  r
   c                 ó   € \        V 4      # ©N)rx   rJ   s   &r	   Únorm_catr–   §   s
   € Ü
˜QÓ
Ðr
   c                ó$   € V ^8„  d   QhR\         /# )r   r    )Úfloat)r   s   "r	   r
   r
   «   s   € ÷ ñ ŒUñ r
   c                 ó^   €  \        V 4      pW8X  d   V# R #   \        \        3 d     R # i ; i)ç        )r˜   Ú	TypeErrorÚ
ValueError)ÚvÚfs   & r	   ÚsvrŸ   «   s6   € ðÜ!‹HˆØ”FˆqÐ# Ð#øÜ”zÐ
"ô Úðús   ‚ • —,«,c                óh   € V ^8„  d   QhR\         \        ,          R\         \        ,          R\        /# )r   Ú recordsÚordered_categoriesr    )r#   Údictr!   )r   s   "r	   r
   r
   ³   s)   € ÷ 8ñ 8œ4¤:ð 8¼4Ä½9ð 8Ìñ 8r
   c                 ó8   € V  Uu0 uF
  q"R ,          kK
  	  ppV Uu. uF
  qDV9   g   K
  VNK  	  pp\        V  Uu0 uF
  q"R,          kK
  	  up4      p\        4       p V  F3  pV P                  VR,          4       V P                  VR,          4       K5  	  \        V 4      pV U	U
u/ uF  p	T	V U
u/ uF  p
V
RRRRR / R/ /bK  	  up
bK!  	  p
p	p
V  EF  pV
P                   VR ,          / 4      P                   VR,          4      pV'       g   K=  VR;;,          VR	,          ,
          uu&   VR;;,          VR
,          ,
          uu&   VR,          P	                  4        F9  w  rÞ\
        VR ,          P                   V
^ 4      V,           ^4      VR ,          V
&   K;  	  VR,          P	                  4        F9  w  rÞ\
        VR,          P                   V
^ 4      V,           ^4      VR,          V
&   K;  	  EK   	  V FG  p	V F>  p
W¹,          V
,          p\
        VR,          ^4      VR&   \
        VR,          ^4      VR&   K@  	  KI  	  / pV  F:  pVP
                  VR,          \        4       4      P                  VR
,          4       K<  	  VP	                  4        U
Uu/ uF  w  p
pV
\        V4      bK  	  pp
pV U	u/ uF  q™/ bK   	  pp	V  F¢  pVR	,          ^ 8”  g   K  VR,           R\        VR
,          4      P                  ^4       2pVP
                  VR ,          / 4       \
        VVR ,          ,          P                   V^ 4      VR	,          ,           ^4      VVR ,          ,          V&   K¤  	  R
VRVR VRV
RVRV RV/ # u upi u upi u upi u up
i u up
p	i u upp
i u up	i )Ú	categoriaÚyearÚ
mxn_ranchesÚ
usd_ranchesÚusdrš   ÚmxnÚ ranchesÚ
ranches_mxnÚ	usd_totalÚ	mxn_totalÚweekz-WÚyearsÚ
categoriesÚ summaryÚweeks_per_yearÚ
weekly_detailÚ
weekly_series)
ÚsortedÚsetÚupdater   ÚitemsÚroundÚ
setdefaultÚaddr!   Úzfill)r¡   r¢   ÚrÚ
cats_foundÚcÚcatsr°   Úranches_seenr«   ÚcatÚyrr²   r3   Úrnr   Údr³   Úwksrµ   Úkeys   &&                  r	   Ú
build_datasetrÉ   ³   st  € Ù*1Ó2©' QK—..©'€JÐ2Ù)Ó
=Ñ)!°*©_AˆAÑ)€DÐ
=Ü¡wÓ/¡w !f—II¡wÑ/Ó0€Eä›€LÛ
ˆØ×Ñ˜A˜mÕ,Ô-Ø×Ñ˜A˜mÕ,Ö-ñ ô \Ó"€Gñ ô
ñ
 ˆCð	 	áó
áð 
˜˜U C¨°B¸
ÀrÐJÒJÙñ
ò 	
ñ ð
 ñ ô ˆØK‰K˜˜+¨Ó+×/Ñ/°°&µ	Ó:ˆßÙØ	ˆ%Ak•NÕ"‹Ø	ˆ%Ak•NÕ"‹Ø}Õ%×+Ñ+Ö-‰EˆBÜ$ Q y¥\×%5Ñ%5°b¸!Ó%<¸qÕ%@À!ÓDˆAˆiL˜Óñ .à}Õ%×+Ñ+Ö-‰EˆBÜ#(¨¨=Õ)9×)=Ñ)=¸bÀ!Ó)DÀqÕ)HÈ!Ó#LˆAˆmÕ˜RÓ ô .ñ ó ˆÛˆBØ•˜RÕ ˆAÜ˜Q˜uX qÓ)ˆAˆe‰HÜ˜Q˜uX qÓ)ˆAˆe‹Hó  ñ ð €NÛ
ˆØ×!Ñ! ! F¥)¬S«UÓ3×7Ñ7¸¸&½	ÖBñ à5C×5IÑ5IÔ5KÔLÑ5K©'¨"¨cbœ& ›+’oÑ5K€NÑLá.2Ó3©d s š7©d€MÐ3Û
ˆØ
ˆ[>˜AÖ
Øv•YK˜r¤# a¨¥i£.×"6Ñ"6°qÓ"9Ð!:Ð;ˆCØ×$Ñ$ Q {¥^°RÔ8Ü16Ø˜a 
nÕ-×1Ñ1°#°qÓ9¸A¸k½NÕJÈAó2ˆM˜!˜K.Õ)¨#Ó.ñ	 ð 	ØdØ7Ø7Ø˜.Ø˜Ø˜ðð ùò_ 3ùÚ
=ùÚ/ùò
ùóùó4 Mùâ3s8   …M7œM<©M<ºNÂ%
N
Â/NÃ N
ÊNÊ2
NÎN
c                ó0   € V ^8„  d   QhR\         R\        /# ©r   Úrowsr    ©r#   r£   )r   s   "r	   r
   r
   ï   s   € ÷ ^ñ ^”Dð ^œTñ ^r
   c                óp   € RRRRRRR RR	R
R
RR
RRRRRRR/
p^p^p^ p^	p/ p/ p V  EFÁ  pV'       d   \        V4      ^
8  d   K  \        V4      V8”  d.   \        W‚,          4      P                  4       P                   4       MRp	\        P
                  ! RRV	4      p	V	'       d   \        V	4      ^8  d   K  \        P                  ! RV	4      '       g   K®  V	R,          p
VP                  V
4      p
V
'       g   V	P                  R4      '       d   Rp
V
'       g   Kó  RV	9   d   RpMRp\        V4      V8”  d    \        Wƒ,          4      P                  4       MRp
V
'       d   V
P                   4       R9   d   EKP  \        V4      V8”  d    \        W„,          4      P                  4       MRp \        \        V4      P                  RR4      4      pV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        V4      V8”  d    \        W…,          4      P                  4       MRp \        \        V4      P                  RR4      4      p\        \        V^4      4      pV'       d   \        V4      MRpV'       d   \        V4      MRpW¼WÙ3pVV 9   d<   V V,          ^ ;;,          V,
          uu&   V V,          ^;;,          V,
          uu&   EKº  VV.V V&   EKÄ  	  V P                  4        F  w  w  r¼rÙw  ppV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        \        V^4      4      pVP                  V
/ 4      P                  V. 4      P!                  V
VVV	.4       K‘  	  V#   \         d     Rp ELœi ; i  \         d     Rp ELEi ; i)zÍ
Lee filas del reporte PR#### del Excel.
  Col 2: UBICACION  (ej: RAMMIPRNN, CECMIPSNF)
  Col 5: PRODUCTO
  Col 7: UNIDADES
  Col 9: GASTO
Retorna: { rancho: { tipo: [[producto, unidades, gasto], ...] } }
ÚVIVr7   ÚRAMrG   ÚISArC   ÚCHRr>   ÚCECrA   ÚC25r?   ÚPOSr9   ÚCAMÚALBr;   ÚHOOr<   r(   rQ   ú
^[A-Z0-9]+$ºNé   NÚMIPÚMIPEÚMIRFEÚ,Ú0rš   ©ÚPRODUCTOÚNOMBREr(   ©Úlenr!   rI   rH   rW   rX   Úmatchr   rw   r˜   r   r"   rº   r   r¹   r»   Úappend©rÌ   Ú	RANCH_MAPÚ
UBICACION_COLÚPRODUCTO_COLÚUNIDADES_COLÚ	GASTO_COLÚresultÚaccumÚrowÚ	ubicacionÚ
ranch_codeÚranchoÚtipoÚproductoÚunidadesÚuÚgastoÚgrÈ   Úu_totÚg_totÚu_strÚg_strs   &                      r	   Ú	_parse_prrþ   ï   s;  € ð 	ˆyØ
ˆzØ
ˆyØ
ˆ{Ø
ˆyØ
ˆ|Ø
ˆzØ
ˆzØ
ˆ}Ø
ˆwð
€Ið €MØ€LØ€LØ€Ià€FØ€Eäˆß”c˜#“h ”mÙä?BÀ3»xÈ-Ô?W”C˜Õ*Ó+×1Ñ1Ó3×9Ñ9Ô;Ð]_ˆ	Ü—F’F˜6 2 yÓ1ˆ	çœC 	›N¨QÔ.ÙÜxŠx˜¨	×2Ò2ÙØ˜r•]ˆ
Ø—‘˜zÓ*ˆ÷ ˜)×.Ñ.¨u×5Ò5ØˆF÷ Ùð
 IÔ
Ø‰DàˆDä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆß˜8Ÿ>™>Ó+Ð/IÔIÚä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆð	Ü”c˜(“m×+Ñ+¨C°Ó4Ó5ˆAØ&'¬3¨q«6¤k”sœ3˜q›6”{´s¼5ÀÀA»;Ó7GˆHô 03°3«x¸)Ô/C”C•NÓ#×)Ñ)Ô+Èˆð	Ü”c˜%“j×(Ñ(¨¨bÓ1Ó2ˆAÜœ˜a ›
Ó$ˆE÷  (ŒE(ŒO¨Sˆß$ŒE%ŒL¨Sˆà˜XÐ1ˆØ
%Œ<Ø#JqM˜QÕ‹MØ#JqM˜QÕŽMà˜Q˜ˆE#ŒJñq ðv @E¿{¹{¾}Ñ;Ñ+ˆx©^¨e°UØ#(¬C°«JÔ#6””C˜“J”¼CÄÀeÈQÃÓ<PˆÜ”E˜% “OÓ$ˆØ×Ñ˜& "Ó%×0Ñ0°°rÓ:×AÑAÀ8ÈUÐTYÐ[dÐBeÖfñ  @Mð
 €Møô5 ô 	Ø‹Hð	ûô ô 	Ø‹Eð	úó%   Æ'ANÈ69N$Î
N!Î N!Î$
N5Î4N5c                ó0   € V ^8„  d   QhR\         R\        /# rË   rÍ   )r   s   "r	   r
   r
   Q  s   € ÷ eñ e”Dð eœTñ er
   c                ód   € RRRRRRR RR	R
R
RR
R/ p^p^p^ p^	p/ p/ p V  EFÁ  pV'       d   \        V4      ^
8  d   K  \        V4      V8”  d.   \        W‚,          4      P                  4       P                   4       MRp	\        P
                  ! RRV	4      p	V	'       d   \        V	4      ^8  d   K  \        P                  ! RV	4      '       g   K®  V	R,          p
VP                  V
4      p
V
'       g   V	P                  R4      '       d   Rp
V
'       g   Kó  RV	9   d   RpMRp\        V4      V8”  d    \        Wƒ,          4      P                  4       MRp
V
'       d   V
P                   4       R9   d   EKP  \        V4      V8”  d    \        W„,          4      P                  4       MRp \        \        V4      P                  RR4      4      pV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        V4      V8”  d    \        W…,          4      P                  4       MRp \        \        V4      P                  RR4      4      p\        \        V^4      4      pV'       d   \        V4      MRpV'       d   \        V4      MRpW¼WÙ3pVV 9   d<   V V,          ^ ;;,          V,
          uu&   V V,          ^;;,          V,
          uu&   EKº  VV.V V&   EKÄ  	  V P                  4        F  w  w  r¼rÙw  ppV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        \        V^4      4      pVP                  V
/ 4      P                  V. 4      P!                  V
VVV	.4       K‘  	  V#   \         d     Rp ELœi ; i  \         d     Rp ELEi ; i)u¢  
Lee filas del reporte MP#### de Google Sheets.
MISMO FORMATO EXACTO que PR####:
  Col 2: UBICACION  (ej: RAMMIPRNN, CECMIPSNF)
  Col 5: PRODUCTO
  Col 7: UNIDADES
  Col 9: GASTO
Retorna: { rancho: { tipo: [[producto, unidades, gasto, ubicacion], ...] } }

Ranchos para MANTENIMIENTO:
  VIV â†’ Prop-RM
  POS â†’ PosCo-RM
  RAM â†’ Campo-RM
  ISA â†’ Isabela
  CEC â†’ Cecilia
  C25 â†’ Cecilia 25
  CHR â†’ Christina
rÏ   r7   rÕ   r9   rÐ   rG   rÑ   rC   rÓ   rA   rÔ   r?   rÒ   r>   r(   rQ   rÙ   rÚ   rÜ   rÝ   rÞ   rß   rà   rš   rá   rä   rè   s   &                      r	   Ú	_parse_mpr  Q  s,  € ð( 	ˆyØ
ˆzØ
ˆzØ
ˆyØ
ˆyØ
ˆ|Ø
ˆ{ð€Ið €MØ€LØ€LØ€Ià€FØ€Eäˆß”c˜#“h ”mÙä?BÀ3»xÈ-Ô?W”C˜Õ*Ó+×1Ñ1Ó3×9Ñ9Ô;Ð]_ˆ	Ü—F’F˜6 2 yÓ1ˆ	çœC 	›N¨QÔ.ÙÜxŠx˜¨	×2Ò2ÙØ˜r•]ˆ
Ø—‘˜zÓ*ˆ÷ ˜)×.Ñ.¨u×5Ò5ØˆF÷ Ùð
 IÔ
Ø‰DàˆDä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆß˜8Ÿ>™>Ó+Ð/IÔIÚä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆð	Ü”c˜(“m×+Ñ+¨C°Ó4Ó5ˆAØ&'¬3¨q«6¤k”sœ3˜q›6”{´s¼5ÀÀA»;Ó7GˆHô 03°3«x¸)Ô/C”C•NÓ#×)Ñ)Ô+Èˆð	Ü”c˜%“j×(Ñ(¨¨bÓ1Ó2ˆAÜœ˜a ›
Ó$ˆE÷  (ŒE(ŒO¨Sˆß$ŒE%ŒL¨Sˆà˜XÐ1ˆØ
%Œ<Ø#JqM˜QÕ‹MØ#JqM˜QÕŽMà˜Q˜ˆE#ŒJñq ðv @E¿{¹{¾}Ñ;Ñ+ˆx©^¨e°UØ#(¬C°«JÔ#6””C˜“J”¼CÄÀeÈQÃÓ<PˆÜ”E˜% “OÓ$ˆØ×Ñ˜& "Ó%×0Ñ0°°rÓ:×AÑAÀ8ÈUÐTYÐ[dÐBeÖfñ  @Mð
 €Møô5 ô 	Ø‹Hð	ûô ô 	Ø‹Eð	ús%   Æ!AN
È09NÎ

NÎNÎ
N/Î.N/c                ó0   € V ^8„  d   QhR\         R\        /# rË   rÍ   )r   s   "r	   r
   r
   º  s   € ÷ dñ d”Dð dœTñ dr
   c                óp   € RRRRRRRR RR	R
R
RR
RRRRRR/
p^p^p^ p^	p/ p/ p V  EFÁ  pV'       d   \        V4      ^
8  d   K  \        V4      V8”  d.   \        W‚,          4      P                  4       P                   4       MRp	\        P
                  ! RRV	4      p	V	'       d   \        V	4      ^8  d   K  \        P                  ! RV	4      '       g   K®  V	R,          p
VP                  V
4      p
V
'       g   V	P                  R4      '       d   Rp
V
'       g   Kó  RV	9   d   RpMRp\        V4      V8”  d    \        Wƒ,          4      P                  4       MRp
V
'       d   V
P                   4       R9   d   EKP  \        V4      V8”  d    \        W„,          4      P                  4       MRp \        \        V4      P                  RR4      4      pV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        V4      V8”  d    \        W…,          4      P                  4       MRp \        \        V4      P                  RR4      4      p\        \        V^4      4      pV'       d   \        V4      MRpV'       d   \        V4      MRpW¼WÙ3pVV 9   d<   V V,          ^ ;;,          V,
          uu&   V V,          ^;;,          V,
          uu&   EKº  VV.V V&   EKÄ  	  V P                  4        F  w  w  r¼rÙw  ppV\        V4      8X  d   \        \        V4      4      M\        \        V^4      4      p\        \        V^4      4      pVP                  V
/ 4      P                  V. 4      P!                  V
VVV	.4       K‘  	  V#   \         d     Rp ELœi ; i  \         d     Rp ELEi ; i)uÖ  
Lee filas del reporte ME#### de Google Sheets.
MISMO FORMATO EXACTO que PR#### y MP####:
  Col 2: UBICACION  (ej: RAMMIRNN, CECMIRSNF)
  Col 5: PRODUCTO
  Col 7: UNIDADES
  Col 9: GASTO
Retorna: { rancho: { tipo: [[producto, unidades, gasto, ubicacion], ...] } }

Ranchos para MATERIAL DE EMPAQUE:
  VIV â†’ Prop-RM
  POS â†’ PosCo-RM
  RAM â†’ Campo-RM
  ISA â†’ Isabela
  CEC â†’ Cecilia
  C25 â†’ Cecilia 25
  CHR â†’ Christina
  ALB â†’ Albahaca-RM
  HOO â†’ HOOPS
rÏ   r7   rÕ   r9   ÚLIMrÐ   rG   rÑ   rC   rÓ   rA   rÔ   r?   rÒ   r>   r×   r;   rØ   r<   r(   rQ   rÙ   rÚ   rÜ   rÝ   rÞ   rß   rà   rš   rá   rä   rè   s   &                      r	   Ú	_parse_mer  º  s5  € ð, 	ˆyØ
ˆzØ
ˆzØ
ˆzØ
ˆyØ
ˆyØ
ˆ|Ø
ˆ{Ø
ˆ}Ø
ˆwð
€Ið €MØ€LØ€LØ€Ià€FØ€Eäˆß”c˜#“h ”mÙä?BÀ3»xÈ-Ô?W”C˜Õ*Ó+×1Ñ1Ó3×9Ñ9Ô;Ð]_ˆ	Ü—F’F˜6 2 yÓ1ˆ	çœC 	›N¨QÔ.ÙÜxŠx˜¨	×2Ò2ÙØ˜r•]ˆ
Ø—‘˜zÓ*ˆç˜)×.Ñ.¨u×5Ò5ØˆFçÙà
IÔ
Ø‰DàˆDä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆß˜8Ÿ>™>Ó+Ð/IÔIÚä58¸³XÀÔ5L”3sÕ(Ó)×/Ñ/Ô1ÐRTˆð	Ü”c˜(“m×+Ñ+¨C°Ó4Ó5ˆAØ&'¬3¨q«6¤k”sœ3˜q›6”{´s¼5ÀÀA»;Ó7GˆHô 03°3«x¸)Ô/C”C•NÓ#×)Ñ)Ô+Èˆð	Ü”c˜%“j×(Ñ(¨¨bÓ1Ó2ˆAÜœ˜a ›
Ó$ˆE÷  (ŒE(ŒO¨Sˆß$ŒE%ŒL¨Sˆà˜XÐ1ˆØ
%Œ<Ø#JqM˜QÕ‹MØ#JqM˜QÕŽMà˜Q˜ˆE#ŒJñg ðj @E¿{¹{¾}Ñ;Ñ+ˆx©^¨e°UØ#(¬C°«JÔ#6””C˜“J”¼CÄÀeÈQÃÓ<PˆÜ”E˜% “OÓ$ˆØ×Ñ˜& "Ó%×0Ñ0°°rÓ:×AÑAÀ8ÈUÐTYÐ[dÐBeÖfñ  @Mð
 €Møô3 ô 	Ø‹Hð	ûô ô 	Ø‹Eð	úrÿ   c                óD   € V ^8„  d   QhR\         P                  R\        /# )r   r   r    )r   r    r£   )r   s   "r	   r
   r
   "  s"   € ÷ Iñ I”r—|‘|ð I¬ñ Ir
   c                 ó–  a5€ . p. p. p. p\        RL4       \        R4       \        RK4       V P                   EFX  pVP                  4       p\        RV R24       VP                   4       \        9   d   \        R4       KI  \
        P                  ! RV\
        P                  4      p\        R\        V4       24       V'       d¿   \
        P                  ! R RV\
        P                  R	7      P                  4       p \        R
V  R24        \        V 4      pR
V^d,          ,           p	\        RV R
V	 RV^d,           24       RT	u;8:  d   R8:  d$   M M \        R4       VP                  WX34       EK<  \        RV	 R24        \
        P                  ! RV\
        P                  4      p
\        R\        V
4       24       V
'       d²   \
        P                  ! RRV\
        P                  R	7      P                  4       p \        V4      p
R
V
^d,          ,           p\        RV
 R
V RV
^d,           24       RTu;8:  d   R8:  d$   M M \        R4       VP                  W]34       EK0  \        RV R24       EKB  V'       d   EKM  \        R4       EK[  	  \        RL4       \        R4       \        R\        V4       24       \        R \        V4       24       \        RM4       V'       g   R!R"/# / pV F  w  pp\        V V^<^#R#7      VV&   K  	  / pR$V UUu. uF  w  ppVNK
  	  upp/pV FO  w  pp\        V VR%^
R#7      p\        V4      pVVV&   V'       d   \!        VP#                  4       4      M. VR&V R'2&   KQ  	  V EF¯  w  pp
VP%                  V. 4      pV'       g   K#  V
^d,          pV
^d,          pR
V,           p\'        R( V 4       ^ R)7      pV Uu. uF$  pVR.V\        V4      ,
          ,          ,           NK&  	  ppRp\        V4      ^8”  d?   \        V^,          4      ^8”  d(   \)        V^,          ^,          4      P                  4       pRNp\+        V4       FD  w  po5\,        ;QJ d    R* S5 4       F  '       g   K
   R+M	  R,M
! R* S5 4       4      '       g   KB  Tp M	  V^ 8  d   EK1  RNp \/        V^,
          \'        ^ V^,
          4      ^,
          RN4       FO  p\,        ;QJ d#    R- VV,           4       F  '       g   K
   R+M	  R,M! R- VV,           4       4      '       g   KM  Tp  M	  V ^ 8  d   EK¼  VV ,          p!\+        V!4       U"U#u. uFE  w  p"p#\1        V#\(        4      '       g   K  V#P                  4       P                   4       R.8X  g   KC  V"NKG  	  p$p"p#V$'       g   EK,  V$^ ,          p%\        V$4      ^8¼  d
   V$^,          MRp&/ / p(p'\+        V!4       Fu  w  p"p#V#'       d   \3        \)        V#4      4      MRp)V)'       g   K.  V"V%8  d   V)V'V"&   K<  V&'       d   V%T"u;8  d   V&8  d
   M M V)V'V"&   K]  V&'       g   Kg  V"V&8”  g   Kp  V)V(V"&   Kw  	  Rp*\/        V^,           \        V4      4       EF´  pVV,          o5\5        V53R/ l\/        ^4       4       R4      p+V+'       g   K6  \7        V+4      p,V,'       d   T,p*KM  \9        V+4      p-\;        V+4      p.V-'       g
   V.'       g   Ku  Rp/Rp0V.'       d
   V*R08w  d   T.p/Tp0MV-'       d
   V*R18w  d   T-p/Tp0MK£  V'P=                  4        U"U)u/ uF+  w  p"p)V"\        S54      8  g   K  V)\?        S5V",          4      bK-  	  p1p"p)V(P=                  4        U"U)u/ uF+  w  p"p)V"\        S54      8  g   K  V)\?        S5V",          4      bK-  	  p2p"p)T0P                  R2T
R3TR4TR5TR6T/R7\A        V%\        S54      8  d   \?        S5V%,          4      M^ ^4      R8\A        V&'       d#   V&\        S54      8  d   \?        S5V&,          4      M^ ^4      R9V1R:V2/	4       EK·  	  EK²  	  \C        V\D        4      p3\C        V\F        4      p4/ R;V3R;,          bR<V3R<,          bR=V3R=,          bR>V3R>,          bR?V3R?,          bR@V3R@,          bRAV3RA,          bRBV4R;,          bRCV4R<,          bRDV4R=,          bREV4R>,          bRFV4R?,          bRGV4R@,          bRHV4RA,          bRIVbRJVb#   \         d   p
\        RT
 24        Rp
?
E L…Rp
?
ii ; i  \         d    \        R4        EKó  i ; iu uppi u upi u up#p"i u up)p"i u up)p"i )OÚ
u!   ðŸ” DETECTANDO HOJAS EN EL EXCELu
   
ðŸ“„ Hoja: 'Ú'u(      â­ï¸  SKIP (en lista de exclusiÃ³n)ú^PR\s*\d{4}$u      ðŸ” Regex PR: úPR\s*r(   ©Úflagsõ      ðŸ“Š CÃ³digo extraÃ­do: 'éÐ   u
      ðŸ“… PRõ
    â†’ AÃ±o ú	, Semana éâ   éî   u      âœ… PR DETECTADA Y VÃLIDAõ      âŒ AÃ±o ú fuera de rango (2018-2030)õ      âŒ Error: Nz^WK\s*\d{4}$u      ðŸ” Regex WK: zWK\s*u
      ðŸ“… WKu      âœ… WK DETECTADA Y VÃLIDAz fuera de rangou!      âŒ Error convirtiendo cÃ³digou      â„¹ï¸  No es WK ni PRu
   ðŸ“Š RESUMEN:u      â€¢ Hojas WK encontradas: u      â€¢ Hojas PR encontradas: Úerrorz#No se encontraron hojas WK validas.)r   r   Úhojas_pr_encontradasiô  ÚPRÚ_ranchosc              3   ó8   "  € T F  p\        V4      x € K  	  R # 5ir•   )rå   )Ú.0r¾   s   & r	   Ú	<genexpr>Ú extraer_datos.<locals>.<genexpr>  s   é € Ð,© 1œ˜AŸ˜«ùs   ‚)Ú defaultc              3   óx   "  € T F0  p\        V\        4      ;'       d    R VP                  4       9   x € K2  	  R# 5i)zEJECUCION SEMANALN)Ú
isinstancer!   rH   )r  rÀ   s   & r	   r  r  ˆ  s0   é € ÐXÑTWÈq”:˜a¤Ó%×JÐJÐ*=ÀÇÁÃÑ*JÔJÓTWùs   ‚:ž:TFc              3   óØ   a"  € T F_  o\        S\        4      ;'       dC    \        ;QJ d#    V3R  l\         4       F  '       g   K
   RM	  RM! V3R  l\         4       4      x € Ka  	  R# 5i)c              3   óH   <"  € T F  qSP                  4       9   x € K  	  R # 5ir•   )rH   )r  Úkr   s   & €r	   r  Ú*extraer_datos.<locals>.<genexpr>.<genexpr>  s   øé € Ð-QÁjÀ°1·7±7³9®nÃjùs   ƒ"TFN)r"  r!   ÚanyÚ
RANCH_KEYS)r  r   s   &@r	   r  r    sA   øé € ÐcÑ[bÐVW”:˜a¤Ó%×QÐQ¯#«#Ô-QÅjÓ-Q¯#¯#ª#Ô-QÅjÓ-QÓ*QÔQÓ[bùs   ƒA*Ÿ
A*«A*Á$A*ÚTOTALc               3   ó
  <"  € T Fx  pV\        S4      8  g   K  SV,          '       g   K&  \        \        SV,          4      P                  4       4      ^8”  g   KV  \        SV,          4      P                  4       x € Kz  	  R# 5i)rÛ   N)rå   r!   rI   )r  rÀ   rð   s   & €r	   r  r  ®  sj   øé € ð W±x°!Ø¤ S£™\ô .Ø.1°!¯f©fô .Ü9<¼SÀÀQÅ»[×=NÑ=NÓ=PÓ9QÐTUÑ9Uô .œ#˜c !f›+×+Ñ+×-Ð-³xùs   ƒBšB«+BÁ(Br\   r]   Úsemanar¦   r¯   Ú
date_ranger¥   r®   r­   r§   r¨   r°   r±   r«   r²   r³   r´   rµ   Úservices_yearsÚservices_categoriesÚservices_ranchesÚservices_summaryÚservices_weeks_per_yearÚservices_weekly_detailÚservices_weekly_seriesÚ	productosÚproductos_debugú<============================================================ú=
============================================================z=============================================================
éÿÿÿÿ)$r   Ú
sheet_namesrI   rH   ÚSKIPrW   ræ   Ú
IGNORECASEÚboolrX   r"   rç   rœ   rå   r0   rþ   r#   Úkeysr   Úmaxr!   Ú	enumerater'  Úranger"  rK   Únextr`   rx   r’   r¹   rŸ   rº   rÉ   ÚMATERIAL_CATEGORIAS_ORDENÚSERVICIOS_CATEGORIAS_ORDEN)6r   Úmateriales_dataÚservicios_dataÚ
hojas_validasÚpr_hojasÚsnameÚpr_matchÚpr_rawÚ pr_codeÚ pr_yearr   Úwk_matchÚcode_rawÚcoder¦   Ú
batch_datar   Ú_r4  Útr5  ÚvalsÚparsedÚrawÚyyÚwwÚmax_colsr¾   Údatar,  Úexec_idxÚiÚ
header_idxr&   Újr   Ú
total_colsÚ
mxn_total_colÚ
usd_total_colÚmxn_ranch_colsÚusd_ranch_colsrÅ   Ú sectionÚlabelÚ
section_foundÚ mat_catÚserv_catÚ
target_catÚ
target_listr§   r¨   Ú
materialesÚ	serviciosrð   s6   &                                                    @r	   Ú
extraer_datosrl  "  sÅ  ø€ Ø€OØ€Nð €MØ€Hä	ˆ/ÔÜ	Ð
-Ô.Ü	ˆ(„Oà—•ˆØ—
‘
“
ˆÜ
˜u˜g QÐ'Ô(à
;‰;‹=œDÔ
 ÜÐ<Ô=Ùô —8’8˜O¨U´B·M±MÓBˆÜ
Ð"¤4¨£>Ð"2Ð3Ô4ß
Ü—V’V˜H b¨%´r·}±}ÔE×KÑKÓMˆFÜÐ0°°¸Ð:Ô;ð

,Ü˜f›+ Ø '¨S¥.Õ1 Ü˜
 7 )¨:°g°Y¸iÈ ÐRUÍ
ÀÐWÔXØ˜7Ö* d×*ÜÐ9Ô:Ø—O‘O UÐ$4Ô5Úä˜L¨¨	Ð1LÐMÕNô
 —8’8˜O¨U´B·M±MÓBˆÜ
Ð"¤4¨£>Ð"2Ð3Ô4ß
Ü—v’v˜h¨¨E¼¿¹ÔG×MÑMÓOˆHð

;Ü˜8“}Ø˜t s{Õ+Ü˜
 4 &¨
°4°&¸	À$ÈÅ*ÀÐNÔOØ˜4Ö' 4×'ÜÐ9Ô:Ø!×(Ñ(¨%¨×7ä˜L¨¨¨oÐ>×?÷ ’8ÜÐ1×2ñ[ !ô^ 
ˆ/ÔÜ	ˆ/ÔÜ	Ð)¬#¨mÓ*<Ð)=Ð
>Ô?Ü	Ð)¬#¨h«-¨Ð
9Ô:Ü	ˆ/Ôç
ØÐ>Ð?Ð?ð €JÛ"‰	ˆÜ'¨¨VÀÐPRÔSˆ
6Óñ #ð €IØ-¹hÔ/G¹h±d°a¸³¹hÒ/GÐH€Oã#‰ˆ Ü˜C °SÀRÔHˆÜ˜4“ˆØ#ˆ	'ÑßIO´$°v·{±{³}Ô2EÐUWˆ˜"˜W˜I XÐ.Ó/ñ	 $ô &‰ˆØn‰n˜V RÓ(ˆßÙàs{ˆØczˆØbyˆäÑ,©Ó,°aÔ8ˆÙ<?Ó@¹C°qA˜˜ ¬3¨q«6Õ 1Õ2×2Ð2¹CˆÐ@àˆ
Ü
ˆt‹9qŒ=œS  a¥›\¨AÔ-Ü˜T !W QZ›×.Ñ.Ó0ˆJàˆÜ –o‰FˆAˆsß‹sÑXÑTWÓXssŠsÑXÑTWÓX×XÔXØÙñ  &ð aŒ<Úàˆ
Üx !•|¤S¨¨H°q­LÓ%9¸AÕ%=¸rÖBˆAß‹sÑcÐ[_Ð`aÖ[bÓcssŠsÑcÐ[_Ð`aÖ[bÓc×cÔcØ
Ùñ  Cð ˜Œ>ÚàjÕ!ˆä$-¨fÔ$5ô NÑ$5™D˜A˜qÜ# A¤s×+ô Ø01· ± ³	·±Ó0AÀWÑ0L÷ aÑ$5ˆ
ñ NçÚØ" 1
ˆ
Ü),¨Z«¸AÔ)=˜
 1ž
À4ˆ
à)+¨R˜ˆÜ˜fÖ%‰DˆAˆqß'(”œC ›FÔ#¨dˆBßÙØ=Ô Ø$&˜qÓ!ß =°1Ö#D°}×#DØ$&˜qÓ!ß‘ 1 }Ö#4Ø$&˜qÓ!ñ &ð ˆ Üx !•|¤S¨£Y×/ˆAØ˜•GˆCÜô W´u¸Q´xó WØX\ó^ˆEçÙä(¨Ó/ˆMßØ' Ùä'¨Ó.ˆGÜ'¨Ó.ˆHß§8ÙàˆJØˆKß˜G {Ô2Ø%
Ø,‘
ß˜W¨
Ô2Ø$
Ø-‘
áà7E×7KÑ7KÔ7MÔ^Ñ7M©e¨a°ÐQRÔUXÐY\ÓU]ÑQ]œ>˜2œr # a¥&›zš>Ñ7MˆKÑ^Ø7E×7KÑ7KÔ7MÔ^Ñ7M©e¨a°ÐQRÔUXÐY\ÓU]ÑQ]œ>˜2œr # a¥&›zš>Ñ7MˆKÑ^à×ÑØ˜tØ˜tØ˜rØ˜zØ˜zØœu¸}ÌsÐSVËxÔ?W¤R¨¨MÕ(:Ô%;Ð]^Ð`aÓbØœu¿}ÐQ^ÔadÐehÓaiÔQi¤R¨¨MÕ(:Ô%;ÐopÐrsÓtØ˜{Ø˜{ð
 ÷ 

ô? 0ñm &ôB ˜Ô0IÓJ€JÜ˜nÔ.HÓI€IðØ ¨GÕ!4ðà ¨LÕ!9ðð 	 ¨IÕ!6ð ð 	 ¨IÕ!6ð	ð
 	 Ð,<Õ!=ð
ð 	 ¨OÕ!<ð
ð 	 ¨OÕ!<ðð 	 ¨7Õ!3ðð 	 ¨<Õ!8ðð 	 ¨9Õ!5ðð 	 ¨9Õ!5ðð 	" 9Ð-=Õ#>ðð 	! 9¨_Õ#=ðð 	! 9¨_Õ#=ðð 	˜9ðð  	˜?ð!ð øôi ô 
,Ü˜ q cÐ*×+Ò+ûð
,ûô" ô 
;ÜÐ9×:Ð:ð
;üó, 0Hùò& Aùó0Nùó` _ùÛ^sn   Ä
A%c)Å3c)Ç?A%dÉ'dÌd.Ï*d4Õd9Õ4 d9Öd9Ü-d?
Ý d?
Ý4e
Þe
ã)
d
ã4däd
äd+ä*d+c                óD   € V ^8„  d   QhR\         R\        P                  /# )r   Úcredentials_pathr    )r!   Ú gspreadÚClient)r   s   "r	   r
   r
   ï  s   € ÷ $ñ $¬ð $ÄgÇnÁnñ $r
   c                 óØ  € ^ RI pRVP                  9   Ed$   RVP                  R,          R,          RVP                  R,          R,          RVP                  R,          R,          RVP                  R,          R,          R VP                  R,          R ,          RVP                  R,          R,          R	VP                  R,          R	,          R
VP                  R,          R
,          R
VP                  R,          R
,          RVP                  R,          R,          /
p\        P                  ! V\        R
7      pM\        P
                  ! V \        R
7      p\        P                  ! V4      # )é    NÚgcp_service_accountÚtypeÚ
project_idÚprivate_key_idÚ
private_keyÚclient_emailÚ	client_idÚauth_uriÚ	token_uriÚauth_provider_x509_cert_urlÚclient_x509_cert_url)Úscopes)Ú	streamlitÚ secretsr   Úfrom_service_account_infoÚSCOPESÚfrom_service_account_filero  Ú	authorize)rn  ÚstÚinfoÚcredss   &   r	   Úget_gsheets_clientrˆ  ï  s  € ÛØ  §
¡
Õ *à¨2¯:©:Ð6KÕ+LÈVÕ+TØ¨2¯:©:Ð6KÕ+LÈ\Õ+ZØ¨2¯:©:Ð6KÕ+LÐM]Õ+^Ø¨2¯:©:Ð6KÕ+LÈ]Õ+[Ø¨2¯:©:Ð6KÕ+LÈ^Õ+\Ø¨2¯:©:Ð6KÕ+LÈ[Õ+YØ¨2¯:©:Ð6KÕ+LÈZÕ+XØ¨2¯:©:Ð6KÕ+LÈ[Õ+YØ)¨2¯:©:Ð6KÕ+LÐMjÕ+kØ"¨2¯:©:Ð6KÕ+LÐMcÕ+dð

ˆô ×5Ò5°dÄ6ÔJ‰ä×5Ò5Ð6FÌvÔVˆÜ
×
Ò
˜UÓ
#Ð#r
   c                óR   € V ^8„  d   QhR\         R\        \        \        3,          /# ©r   Úspreadsheet_namer    ©r!   Útupler£   )r   s   "r	   r
   r
     s(   € ÷ G&ñ G&¬Sð G&ÄEÌ$ÔPTÈ*ÕDUñ G&r
   c                óv  € / pR. /p \        4       pRpW P                  RR4      3 F  p VP                  V4      p M	  VfN   VP
                  4        F9  pRVP                  P                  4       9   g   K$  RVP                  9   g   K7  Tp M	  Vf   \        R 4       W3# . p VP                  4        Fß  pVP                  P                  4       p	\        P                  ! RV	\        P                  4      p
V
'       g   KM  \        P                  ! R	R
V	\        P                  R
7      P                  4       p
 \        V
4      pRV^d,          ,           p
R
T
u;8:  d   R8:  d2   M K²  \        RV	 24       V P!                  VP                  V34       Kß  Ká  	  V  UUu. uF   w  rïVNK	  	  uppVR&   V '       g   \        R4       W3# ^dpV  UUu. uF
  w  rïRV R2NK
  	  ppp\%        ^ \'        V4      V4       FØ  pVVVV,            pVP)                  VRR/R7      pVP+                  R. 4       Fž  pVP+                  RR
4      pVP+                  R. 4      pVP-                  R4      ^ ,          P                  R4      pV  FJ  w  ppVV8X  g   K  \/        V4      pVVV&   V'       d   \1        VP3                  4       4      M. VRV R2&    Kœ  	  K   	  KÚ  	  W3#   \        P                   d     EKÉ  i ; i  \"         d     EKN  i ; iu uppi u uppi   \4         d   p\        RT 24        Rp?Y3# Rp?ii ; i)z†
Se conecta a Google Sheets y lee SOLO las hojas PR####.
Retorna (productos, productos_debug) con el mismo formato que extraer_datos.
r  NrR   rQ  ÚWKÚ2026u5   âš ï¸  No se encontrÃ³ el Google Sheet para hojas PRr
  r  r(   r
  r  r  r  u       âœ… PR encontrada en Sheets: u-      â„¹ï¸  No hay hojas PR en el Google Sheetr
  ú	'!A1:K500ÚvalueRenderOptionÚUNFORMATTED_VALUE©ÚparamsÚ
valueRangesr@  r-   Ú!r  r  u.   âš ï¸  Error leyendo PR desde Google Sheets: )rˆ  r   Úopenro  ÚSpreadsheetNotFoundÚ openallÚtitlerH   r   Ú
worksheetsrI   rW   ræ   r;  rX   r"   rç   rœ   r@  rå   Úvalues_batch_getr   Úsplitrþ   r#   r=  r   )r‹  r4  r5  ÚclientÚssÚnamer3   rG  ÚwsrH  rI  rJ  rK  rL  rR  rQ  ÚBATCHÚ	pr_rangosr[  ÚgrupoÚresÚitemÚrngrS  ÚtitÚptÚpcrT  r   s   &                            r	   Ú_fetch_pr_desde_sheetsr¬    s  € ð
 €IØ-¨rÐ2€Oð=DÜ#Ó%ˆà
ˆØ%×'?Ñ'?ÀÀSÓ'IÓJˆDð
Ø—[‘[ Ó&Ùñ  Kð Š:Ø—^‘^Ö%Ø˜1Ÿ7™7Ÿ=™=›?Ö*¨v¸¿¹Ö/@ØBÙñ  &ð
 Š:ÜÐIÔJØÐ-Ð-ð ˆØ—-‘-–/ˆBØ—H‘H—N‘NÓ$ˆEÜ—x’x °¼¿
¹
ÓFˆHß‰xÜŸš ¨"¨e¼2¿=¹=ÔI×OÑOÓQð Ü! &›kGØ" g°¥nÕ5GØ˜wÖ.¨$×.Ð.ÜÐ @ÀÀ ÐHÔIØ Ÿ™¨¯©°7Ð(;Ö<ñ /ñ "ñ BJÔ2JÁ¹¸³1ÁÒ2JˆÐ.Ñ/çÜÐAÔBØÐ-Ð-ð ˆÙ2:Ô;±(©$¨!q˜˜˜9Ó%±(ˆ	Ñ;Üqœ#˜i›.¨%Ö0ˆAØ˜a  E¥	Ð*ˆEØ×'Ñ'¨Ð7JÐL_Ð6`Ð'ÓaˆCØŸ ™  
¨rÖ2Ø—x‘x  ¨Ó,Ø—x‘x ¨"Ó-Ø—y‘y “~ aÕ(×.Ñ.¨sÓ3Û&‘FB˜Ø˜S–yÜ!*¨4£˜Ø(.˜	 "™
ßTZ¼TÀ&Ç+Á+Ã-Ô=PÐ`b˜¨"¨R¨D°Ð(9Ñ:Úó
 'ó	 3ñ  1ð" Ð
%Ð%øôo ×.Ñ.ô 
Ûð
ûô4 "ô Ûðüó 3Kùó <øô ô DÜ
Ð>¸q¸cÐB×CÐCà
Ð
%Ð%ûð Dús®   ˆ"L «K¼7L Á8L Â
L Â#AL Ä6L Ä9,K4Å%L Å'+K4Æ
L Æ
L Æ*L Æ9
L Ç  L ÇL
ÇB)L Ê
A L ËK1Ë,L Ë0K1Ë1L Ë4
LË?L ÌLÌL Ì
L8ÌL3Ì3L8c                óR   € V ^8„  d   QhR\         R\        \        \        3,          /# rŠ  rŒ  )r   s   "r	   r
   r
   P  ó(   € ÷ M,ñ M,¬Sð M,ÄEÌ$ÔPTÈ*ÕDUñ M,r
   c                ón   € / pR. /p \        4       pRpW P                  RR4      3 F  p VP                  V4      p M	  VfN   VP
                  4        F9  pRVP                  P                  4       9   g   K$  RVP                  9   g   K7  Tp M	  Vf   \        R 4       W3# . p VP                  4        EF  pVP                  P                  4       p	\        P                  ! RV	\        P                  4      p
V
'       g   KN  \        P                  ! R	R
V	\        P                  R
7      P                  4       p
\        RV
 R
24        \        V
4      pRV^d,          ,           p
\        RV RV
 RV^d,           24       RT
u;8:  d   R8:  d2   M M.\        RV	 24       V P!                  VP                  V34       EK
  \        RV
 R24       EK  	  V  UUu. uF  w  ppVNK
  	  uppVR&   V '       g   \        R4       W3# ^dpV  UUu. uF  w  ppR
V R2NK  	  ppp\%        ^ \'        V4      V4       EF  pVVVV,            pVP)                  VRR/R7      pVP+                  R. 4       FÆ  pVP+                  RR
4      pVP+                  R. 4      pVP-                  R 4      ^ ,          P                  R
4      pV  Fr  w  ppVV8X  g   K  \/        V4      pVVV&   V'       d   \1        VP3                  4       4      M. VR!V R"2&   \        R#V R$\1        VP3                  4       4       24        KÄ  	  KÈ  	  EK  	  W3#   \        P                   d     EK0  i ; i  \"         d   p\        RT 24        Rp?EKÆ  Rp?ii ; iu uppi u uppi   \4         d   p\        R%T 24        Rp?Y3# Rp?ii ; i)&u«   
Se conecta a Google Sheets y lee SOLO las hojas MP####.
PatrÃ³n: MP2611 â†’ aÃ±o 2026, semana 11.
Retorna (productos_mp, productos_mp_debug) con el mismo formato que PR.
Úhojas_mp_encontradasNrR   rQ  r  r  u5   âš ï¸  No se encontrÃ³ el Google Sheet para hojas MPz^MP\s*\d{4}$zMP\s*r(   r
  r  r
  r  u
      ðŸ“… MPr  r  r  r  u       âœ… MP encontrada en Sheets: r  r  r  u-      â„¹ï¸  No hay hojas MP en el Google Sheetr‘  r’  r“  r”  r–  r@  r-   r—  ÚMPr  u
      ðŸ„ MPú ranchos detectados: u.   âš ï¸  Error leyendo MP desde Google Sheets: )rˆ  r   r˜  ro  r™  rš  r›  rH   r   rœ  rI   rW   ræ   r;  rX   r"   rç   rœ   r@  rå   r  r   rž  r  r#   r=  r   )r‹  Úproductos_mpÚproductos_mp_debugrŸ  r   r¡  r3   Úmp_hojasr¢  rH  Úmp_matchÚmp_rawÚ mp_codeÚ mp_yearr   rR  rQ  r£  Ú	mp_rangosr[  r¥  r¦  r§  r¨  rS  r©  rª  r«  rT  s   &                            r	   Ú_fetch_mp_desde_sheetsr»  P  ó  € ð €LØ0°"Ð5ÐðBDÜ#Ó%ˆà
ˆØ%×'?Ñ'?ÀÀSÓ'IÓJˆDð
Ø—[‘[ Ó&Ùñ  Kð Š:Ø—^‘^Ö%Ø˜1Ÿ7™7Ÿ=™=›?Ö*¨v¸¿¹Ö/@ØBÙñ  &ð
 Š:ÜÐIÔJØÐ3Ð3ð ˆØ—-‘-—/ˆBØ—H‘H—N‘NÓ$ˆEÜ—x’x °¼¿
¹
ÓFˆHß‰xÜŸš ¨"¨e¼2¿=¹=ÔI×OÑOÓQÜÐ4°V°H¸AÐ>Ô?ð
0Ü! &›kGØ" g°¥nÕ5GÜ˜J w i¨z¸'¸À)ÈGÐVYÍMÈ?Ð[Ô\Ø˜wÖ.¨$×.ÜÐ @ÀÀ ÐHÔIØ Ÿ™¨¯©°7Ð(;×<ä ¨W¨IÐ5PÐQ×Rñ "ñ$ EMÔ5MÁH¹D¸A¸q³aÁHÒ5MÐÐ1Ñ2çÜÐAÔBØÐ3Ð3ð ˆÙ2:Ô;±(©$¨!¨Qq˜˜˜9Ó%±(ˆ	Ñ;Üqœ#˜i›.¨%×0ˆAØ˜a  E¥	Ð*ˆEØ×'Ñ'¨Ð7JÐL_Ð6`Ð'ÓaˆCØŸ ™  
¨rÖ2Ø—x‘x  ¨Ó,Ø—x‘x ¨"Ó-Ø—y‘y “~ aÕ(×.Ñ.¨sÓ3Û&‘FB˜Ø˜S–yÜ!*¨4£˜Ø+1˜ RÑ(ßW]ÄÀVÇ[Á[Ã]Ô@SÐceÐ*¨R°¨t°8Ð+<Ñ=Ü 
¨2¨$Ð.CÄDÈÏÉËÓDWÐCXÐYÔZÚó
 'ô	 3ñ  1ð$ Ð
+Ð+øôy ×.Ñ.ô 
Ûð
ûô< "ô 0Ü˜N¨1¨#Ð.×/Ó/ûð0üó 6Nùó <øô  ô DÜ
Ð>¸q¸cÐB×CÐCà
Ð
+Ð+ûð Dúó¶   ˆ"N «L>¼7N Á8N Â
N Â#AN ÄAN Å	A3MÆ<N Æ?MÇ
N ÇNÇ&N Ç5
N È N È
N	ÈB*N Ë
A0N Ì>MÍN ÍMÍN Í
N Í&M;Í4 N Í;N Î N Î
N4ÎN/Î/N4c                óR   € V ^8„  d   QhR\         R\        \        \        3,          /# rŠ  rŒ  )r   s   "r	   r
   r
   ¡  r®  r
   c                ón   € / pR. /p \        4       pRpW P                  RR4      3 F  p VP                  V4      p M	  VfN   VP
                  4        F9  pRVP                  P                  4       9   g   K$  RVP                  9   g   K7  Tp M	  Vf   \        R 4       W3# . p VP                  4        EF  pVP                  P                  4       p	\        P                  ! RV	\        P                  4      p
V
'       g   KN  \        P                  ! R	R
V	\        P                  R
7      P                  4       p
\        RV
 R
24        \        V
4      pRV^d,          ,           p
\        RV RV
 RV^d,           24       RT
u;8:  d   R8:  d2   M M.\        RV	 24       V P!                  VP                  V34       EK
  \        RV
 R24       EK  	  V  UUu. uF  w  ppVNK
  	  uppVR&   V '       g   \        R4       W3# ^dpV  UUu. uF  w  ppR
V R2NK  	  ppp\%        ^ \'        V4      V4       EF  pVVVV,            pVP)                  VRR/R7      pVP+                  R. 4       FÆ  pVP+                  RR
4      pVP+                  R. 4      pVP-                  R 4      ^ ,          P                  R
4      pV  Fr  w  ppVV8X  g   K  \/        V4      pVVV&   V'       d   \1        VP3                  4       4      M. VR!V R"2&   \        R#V R$\1        VP3                  4       4       24        KÄ  	  KÈ  	  EK  	  W3#   \        P                   d     EK0  i ; i  \"         d   p\        RT 24        Rp?EKÆ  Rp?ii ; iu uppi u uppi   \4         d   p\        R%T 24        Rp?Y3# Rp?ii ; i)&u°   
Se conecta a Google Sheets y lee SOLO las hojas ME####.
PatrÃ³n: ME2611 â†’ aÃ±o 2026, semana 11.
Retorna (productos_me, productos_me_debug) con el mismo formato que PR y MP.
Úhojas_me_encontradasNrR   rQ  r  r  u5   âš ï¸  No se encontrÃ³ el Google Sheet para hojas MEz^ME\s*\d{4}$zME\s*r(   r
  r  r
  r  u
      ðŸ“… MEr  r  r  r  u       âœ… ME encontrada en Sheets: r  r  r  u-      â„¹ï¸  No hay hojas ME en el Google Sheetr‘  r’  r“  r”  r–  r@  r-   r—  ÚMEr  u
      ðŸ“¦ MEr²  u.   âš ï¸  Error leyendo ME desde Google Sheets: )rˆ  r   r˜  ro  r™  rš  r›  rH   r   rœ  rI   rW   ræ   r;  rX   r"   rç   rœ   r@  rå   r  r   rž  r  r#   r=  r   )r‹  Úproductos_meÚproductos_me_debugrŸ  r   r¡  r3   Úme_hojasr¢  rH  Úme_matchÚme_rawÚ me_codeÚ me_yearr   rR  rQ  r£  Ú	me_rangosr[  r¥  r¦  r§  r¨  rS  r©  rª  r«  rT  s   &                            r	   Ú_fetch_me_desde_sheetsrÊ  ¡  r¼  r½  c                ó0   € V ^8„  d   QhR\         R\        /# rŠ  )r!   r£   )r   s   "r	   r
   r
   ò  s   € ÷ ,ñ ,¤ð ,´tñ ,r
   c                óþ  € \        4       pVf   RR/#  \        P                  ! V4      p\	        T4      pRT9  d£   \
        R4       \
        R4       \
        R4       \
        T 4      w  rVYTR&   YdR &   \
        R4       \
        R4       \
        R4       \        T 4      w  rxYtR	&   Y„R
&   \
        R4       \
        R
4       \
        R4       \        T 4      w  ršY”R&   Y¤R
&   T#   \         d   pRRT 2/u Rp?# Rp?ii ; i)uø   
- Hojas WK  â†’ descargadas desde el Excel de OneDrive
- Hojas PR  â†’ leÃ­das desde Google Sheets (productos generales)
- Hojas MP  â†’ leÃ­das desde Google Sheets (MANTENIMIENTO)
- Hojas ME  â†’ leÃ­das desde Google Sheets (MATERIAL DE EMPAQUE)
Nr  z,No se pudo descargar el archivo de OneDrive.zNo se pudo abrir el Excel: u)   ðŸ” LEYENDO HOJAS PR DESDE GOOGLE SHEETSr4  r5  u9   ðŸ” LEYENDO HOJAS MP DESDE GOOGLE SHEETS (MANTENIMIENTO)r³  r´  u?   ðŸ” LEYENDO HOJAS ME DESDE GOOGLE SHEETS (MATERIAL DE EMPAQUE)rÂ  rÃ  r6  r7  )	r   r   r    r   rl  r   r¬  r»  rÊ  )
r‹  Ú archivor   r   Ú	resultador4  r5  r³  r´  rÂ  rÃ  s
   &          r	   Ú	get_datosrÏ  ò  s  € ô Ó€GØ ‚ØÐGÐHÐHð<ÜlŠl˜7Ó#ˆô ˜cÓ"€Ià iÔ ä
ˆoÔÜ
Ð9Ô:Ü
ˆhŒÜ%;Ð<LÓ%MÑ"ˆ	Ø'0+ÑØ'6Ð#Ñ$ô 	ˆoÔÜ
ÐIÔJÜ
ˆhŒÜ+AÐBRÓ+SÑ(ˆØ*6.Ñ!Ø*<Ð&Ñ'ô 	ˆoÔÜ
ÐOÔPÜ
ˆhŒÜ+AÐBRÓ+SÑ(ˆØ*6.Ñ!Ø*<Ð&Ñ'à
Ðøô= ô <ØÐ6°q°cÐ:Ð;Õ;ûð<ús   ”C  Ã 
C<Ã+C7Ã1C<Ã7C<)r6   r8   rD   rB   r<   r@   r=   r:   )
re   rf   rh   ri   rk   rm   rp   rq   rs   ru   rv   )r|   r   r   rƒ   r†   rŠ   rŒ   r‘   >   ÚDATOSÚHOJA1ÚSHEET1Ú	ACUMULADOÚ
COMPARATIVOú
GRAFICOS I-IV)é<   é#   )zcredentials.json)z
WK 2026-08)$Ú __doc__rW   rS   r   Úpandasr   ro  Úior   Úgoogle.oauth2.service_accountr   r‚  r   r(  rB  rC  r:  r   r0   rK   rY   r`   rx   r’   r–   rŸ   rÉ   rþ   r  r  rl  rˆ  r¬  r»  rÊ  rÏ  © r
   r	   Ú<module>rÝ     sÎ   ðñó 
Û Û Û Û Ý Ý 5ð <Ø4ð
€ðRð 
ò _€
òÐ ò	Ð ò Q€õ
÷õ0õ
õõ
õ õ* õõ8õx^õDeõRdõPI÷Z$÷,G&÷VM,÷bM,÷b,ñ ,r
   
