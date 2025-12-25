import pandas as pd
import numpy as np
import yfinance as yf
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.model_selection import train_test_split
from sklearn.utils import compute_sample_weight
from sklearn.metrics import accuracy_score, precision_score, recall_score, roc_auc_score
from xgboost import XGBClassifier
from imblearn.over_sampling import SMOTE
from sklearn.preprocessing import StandardScaler
import warnings
import time
from datetime import datetime, timedelta
from collections import Counter
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from datetime import datetime
from zoneinfo import ZoneInfo  # Python 3.9+

# ==========================================
# CONFIGURATION & STATE MANAGEMENT
# ==========================================
LOG_FILE = "trading_log.xlsx"
IST = ZoneInfo("Asia/Kolkata")

# Define the list of NSE tickers (example subset for testing)
nse_tickers = [  'HCLTECH.NS', 'OLAELEC.NS', 'BSE.NS', 'GODFRYPHLP.NS', 'HDFCBANK.NS', 'JPPOWER.NS', 'RALLIS.NS', 'INFY.NS', 'WAAREERTL.NS', 'FACT.NS', 'NEULANDLAB.NS', 'ETERNAL.NS', 'ANANTRAJ.NS', 'MIDHANI.NS', 'M&M.NS', 'RELIANCE.NS', 'TCS.NS', 'SBIN.NS', 'HEROMOTOCO.NS', 'DMART.NS', 'BHARTIARTL.NS', 'WAAREEENER.NS', 'KOTAKBANK.NS', 'ICICIBANK.NS', 'PEL.NS', 'SUNPHARMA.NS', 'TATATECH.NS', 'LT.NS', 'AXISBANK.NS', 'TATAELXSI.NS', 'SWIGGY.NS', 'TEJASNET.NS', 'BEL.NS', 'PAYTM.NS', 'AARTIDRUGS.NS', 'AMBER.NS', 'TATAMOTORS.NS', 'CDSL.NS', 'GRSE.NS', 'PGEL.NS', 'SUZLON.NS', 'JIOFIN.NS', 'TECHM.NS', 'EICHERMOT.NS', 'PPLPHARMA.NS', 'HAL.NS', 'HDFCAMC.NS', 'HSCL.NS', 'GLENMARK.NS', 'CHOLAFIN.NS', 'INDIGO.NS', 'MCX.NS', 'RVNL.NS', 'ITC.NS', 'BAJFINANCE.NS', 'HINDUNILVR.NS', 'APOLLOHOSP.NS', 'MAZDOCK.NS', 'PRESTIGE.NS', 'INDUSINDBK.NS', 'CUMMINSIND.NS', 'SOBHA.NS', 'DIXON.NS', 'COFORGE.NS', 'HINDALCO.NS', 'CRIZAC.NS', 'BIOCON.NS', 'HINDPETRO.NS', 'PERSISTENT.NS', 'YESBANK.NS', 'POLYCAB.NS', 'BANDHANBNK.NS', 'TRENT.NS', 'BALKRISIND.NS', 'CANBK.NS', 'BOSCHLTD.NS', 'TVSMOTOR.NS', 'BAJAJ-AUTO.NS', 'RAILTEL.NS', 'AIIL.NS', 'SWARAJENG.NS', 'IDEA.NS', 'NTPC.NS', 'TATASTEEL.NS', 'ANGELONE.NS', 'MOTHERSON.NS', 'ABB.NS', 'BDL.NS', 'EIEL.NS', 'SONACOMS.NS', 'DEEPAKFERT.NS', 'VBL.NS', 'INDHOTEL.NS', 'BEML.NS', 'ICICIPRULI.NS', 'WIPRO.NS', 'PNB.NS', 'VEDL.NS', 'ADANIGREEN.NS', 'ADANIPOWER.NS', 'HDFCLIFE.NS', 'PATANJALI.NS', 'TITAN.NS', 'IDFCFIRSTB.NS', 'ENRIN.NS', 'NAUKRI.NS', 'SOLARINDS.NS', 'SBILIFE.NS', 'SHRIRAMFIN.NS', 'CAMS.NS', 'DLF.NS', 'VOLTAS.NS', 'TEGA.NS', 'TRANSRAILL.NS', 'ZEEL.NS', 'ANANDRATHI.NS', 'COCHINSHIP.NS', 'BELRISE.NS', 'LUPIN.NS', 'RTNPOWER.NS', 'GODREJPROP.NS', 'BANKBARODA.NS', 'SILVERBEES.NS', 'FORTIS.NS', 'AUBANK.NS', 'NMDC.NS', 'POWERGRID.NS', 'BAJAJFINSV.NS', 'ONGC.NS', 'BPCL.NS', 'GILLETTE.NS', 'LAURUSLABS.NS', 'NIFTYBEES.NS', 'WOCKPHARMA.NS', 'AEROFLEX.NS', 'CIPLA.NS', 'POWERINDIA.NS', 'FORCEMOT.NS', 'VMM.NS', 'PREMIERENE.NS', 'ASHAPURMIN.NS', 'ULTRACEMCO.NS', 'GLAND.NS', 'CHENNPETRO.NS', 'COALINDIA.NS', 'LICHSGFIN.NS', 'MAHABANK.NS', 'PFC.NS', 'NBCC.NS', 'SBICARD.NS', 'ICICIGI.NS', 'MARUTI.NS', 'AADHARHFC.NS', 'NATCOPHARM.NS', 'KALYANKJIL.NS', 'ASHOKLEY.NS', 'SAGILITY.NS', 'KPITTECH.NS', 'BHARATFORG.NS', 'RBLBANK.NS', 'KEI.NS', 'BLUESTARCO.NS', 'MANKIND.NS', 'DEVYANI.NS', 'DRREDDY.NS', 'MRF.NS', 'IRFC.NS', 'MOBIKWIK.NS', 'JSWENERGY.NS', 'HEXT.NS', 'TATAPOWER.NS', 'STLTECH.NS', 'TORNTPHARM.NS', 'UNIONBANK.NS', 'GAIL.NS', 'COLPAL.NS', 'CUPID.NS', 'GRAPHITE.NS', 'ASIANPAINT.NS', 'RECLTD.NS', 'SYNGENE.NS', 'IREDA.NS', 'INOXWIND.NS', 'VIPIND.NS', 'PCJEWELLER.NS', 'UPL.NS', 'TRAVELFOOD.NS', 'CEATLTD.NS', 'NAZARA.NS', 'PETRONET.NS', 'FEDERALBNK.NS', 'DBREALTY.NS', 'PARADEEP.NS', 'OBEROIRLTY.NS', 'MFSL.NS', 'COROMANDEL.NS', 'GREAVESCOT.NS', 'HUDCO.NS', 'UNITDSPR.NS', 'NAM-INDIA.NS', 'SYRMA.NS', 'JMFINANCIL.NS', 'CROMPTON.NS', 'KPRMILL.NS', 'HINDZINC.NS', 'INDUSTOWER.NS', 'AJANTPHARM.NS', 'GMRAIRPORT.NS', 'NHPC.NS', 'SAMMAANCAP.NS', 'DIACABS.NS', 'LTF.NS', 'ASTRAZEN.NS', 'NESTLEIND.NS', 'KAYNES.NS', 'CGPOWER.NS', 'LODHA.NS', 'KARURVYSYA.NS', 'HAVELLS.NS', 'NESCO.NS', 'SIEMENS.NS', 'RCF.NS', 'KPIGREEN.NS', 'MPHASIS.NS', 'JSWSTEEL.NS', 'PCBL.NS', 'ADANIENT.NS', 'DCMSHRIRAM.NS', 'UJJIVANSFB.NS', 'MAXHEALTH.NS', 'JINDALSTEL.NS', 'SPANDANA.NS', 'ADANIPORTS.NS', 'MCLOUD.NS', 'KRN.NS', 'BLS.NS', 'NYKAA.NS', 'SIGNATURE.NS', 'SWSOLAR.NS', 'DELHIVERY.NS', 'AUROPHARMA.NS', 'SRF.NS', 'FIVESTAR.NS', 'NUVOCO.NS', 'ALKEM.NS', 'AMBUJACEM.NS',
              'ROSSARI.NS', 'GODREJCP.NS', 'ABDL.NS', 'CARTRADE.NS', 'PNBHOUSING.NS', 'KAJARIACER.NS', 'LTIM.NS', 'ASTERDM.NS', 'KIRLOSBROS.NS', 'HEG.NS', 'QUADFUTURE.NS', 'LTFOODS.NS', 'AWL.NS', 'POLICYBZR.NS', 'APLAPOLLO.NS', 'PVRINOX.NS', 'POWERMECH.NS', 'KFINTECH.NS', 'ALLCARGO.NS', 'DIVISLAB.NS', 'MANAPPURAM.NS', 'CHAMBLFERT.NS', 'HYUNDAI.NS', 'GOLDBEES.NS', 'TRITURBINE.NS', 'NCC.NS', 'MSUMI.NS', 'IOC.NS', 'OFSS.NS', 'CONCOR.NS', 'MOTILALOFS.NS', 'GRASIM.NS', 'IDBI.NS', 'SMLISUZU.NS', 'WEBELSOLAR.NS', 'GABRIEL.NS', 'AEGISVOPAK.NS', 'TITAGARH.NS', 'BHEL.NS', 'JSWINFRA.NS', 'RKFORGE.NS', 'INDIANB.NS', 'BAJAJHCARE.NS', 'AHLUCONT.NS', 'BANKINDIA.NS', 'MUTHOOTFIN.NS', 'CYIENT.NS', 'VINCOFE.NS', 'OIL.NS', 'BRITANNIA.NS', 'JBCHEPHARM.NS', 'IFCI.NS', 'COHANCE.NS', 'NTPCGREEN.NS', 'TIMKEN.NS', 'JAYNECOIND.NS', 'NH.NS', 'CESC.NS', 'ADANIENSOL.NS', 'WELCORP.NS', 'TIINDIA.NS', 'LALPATHLAB.NS', 'JGCHEM.NS', 'FIRSTCRY.NS', 'SAIL.NS', 'DABUR.NS', 'MANORAMA.NS', 'IIFL.NS', 'RAJOOENG.NS', 'IPCALAB.NS', 'SUPREMEIND.NS', 'SAMHI.NS', 'GMRP&UI.NS', 'AARTIIND.NS', 'PHOENIXLTD.NS', 'IRCON.NS', 'ACMESOLAR.NS', 'LIQUIDCASE.NS', 'MARICO.NS', 'SPMLINFRA.NS', 'DOMS.NS', 'SJVN.NS', 'ABCAPITAL.NS', 'PTC.NS', 'NETWEB.NS', 'VPRPL.NS', 'MAMATA.NS', 'IEX.NS', 'OSWALPUMPS.NS', 'TATACONSUM.NS', 'TANLA.NS', 'ASTRAL.NS', 'RADICO.NS', 'NUVAMA.NS', 'SUPRIYA.NS', 'SHAKTIPUMP.NS', 'BSOFT.NS', 'J&KBANK.NS', 'BAJAJHFL.NS', 'BAYERCROP.NS', 'AETHER.NS', 'NATIONALUM.NS', 'ENGINERSIN.NS', 'ITCHOTELS.NS', 'DCBBANK.NS', 'ELLEN.NS', 'GENUSPOWER.NS', 'RELIGARE.NS', 'PAGEIND.NS', 'ITBEES.NS', 'BAJAJHLDNG.NS', 'ZYDUSLIFE.NS', 'TIMETECHNO.NS', 'IRB.NS', 'PIIND.NS', 'ABFRL.NS', 'DIFFNKG.NS', 'PROTEAN.NS', 'NFL.NS', 'TFCILTD.NS', 'KPEL.NS', 'BOMDYEING.NS', 'JUSTDIAL.NS', 'INDGN.NS', 'KIRLOSENG.NS', 'LEMONTREE.NS', 'RAJESHEXPO.NS', 'JKCEMENT.NS', 'CGCL.NS', 'HFCL.NS', 'BRIGADE.NS', 'NORTHARC.NS', 'DATAPATTNS.NS', 'JUBLINGREA.NS', 'IGL.NS', 'APARINDS.NS', '360ONE.NS', 'TATACHEM.NS', 'PFOCUS.NS', 'GRAVITA.NS', 'AGIIL.NS', 'MTNL.NS', 'APOLLOTYRE.NS', 'HPL.NS', 'CERA.NS', 'QPOWER.NS', 'ONESOURCE.NS', 'AARTIPHARM.NS', 'CASTROLIND.NS', 'LAXMIDENTL.NS', 'ACUTAAS.NS', 'SUMICHEM.NS', 'HOMEFIRST.NS', 'STAR.NS', 'M&MFIN.NS', 'USHAMART.NS', 'SOUTHBANK.NS', 'SIGACHI.NS', 'HDFCSML250.NS', 'CHOICEIN.NS', 'APTUS.NS', 'TEXRAIL.NS', 'JYOTICNC.NS', 'VARROC.NS', 'HINDCOPPER.NS', 'GMDCLTD.NS', 'PRAJIND.NS', 'LICI.NS', 'BLUEJET.NS', 'WENDT.NS', 'PARAS.NS', 'VOLTAMP.NS', 'HBLENGINE.NS', 'TARIL.NS', 'RATEGAIN.NS', 'REDINGTON.NS', 'ABSLAMC.NS', 'PROSTARM.NS', 'SONATSOFTW.NS', 'PENIND.NS', 'ZENSARTECH.NS', 'LLOYDSENT.NS', 'JKLAKSHMI.NS', 'DHANUKA.NS', 'EXIDEIND.NS', 'ELECON.NS', 'STARHEALTH.NS', 'SAILIFE.NS', 'ABLBL.NS', 'POCL.NS', 'MRPL.NS', 'LTTS.NS', 'SERVOTECH.NS', 'ARKADE.NS', 'GARUDA.NS', 'RAYMONDREL.NS', 'FSL.NS', 'MARATHON.NS', 'BHARTIHEXA.NS', 'SHARDACROP.NS', 'GICRE.NS', 'DMCC.NS', 'METROPOLIS.NS', 'MEDANTA.NS', 'RUPA.NS', 'SENCO.NS', 'SHREECEM.NS', 'EDELWEISS.NS', 'IGARASHI.NS', 'CRAFTSMAN.NS', 'KEC.NS', 'TBOTEK.NS', 'LLOYDSENGG.NS', 'RAMCOCEM.NS', 'ATGL.NS', 'EMBDL.NS', 'PARAGMILK.NS', 'TECHNOE.NS', 'ACC.NS', 'MBAPL.NS', 'PSPPROJECT.NS', 'IRCTC.NS', 'KIRLOSIND.NS', 'BLACKBUCK.NS', 'DATAMATICS.NS', 'UNOMINDA.NS', 'YATHARTH.NS', 'OLECTRA.NS', 'UTIAMC.NS', 'ELECTCAST.NS', 'RTNINDIA.NS', 'EQUITASBNK.NS', 'IGIL.NS', 'APOLLO.NS', 'RAYMOND.NS', 'ARE&M.NS', 'PIDILITIND.NS', 'PAISALO.NS', 'TVSHLTD.NS', 'KIMS.NS', 'HDFCSILVER.NS', 'UCOBANK.NS', 'ESCORTS.NS', 'MASTEK.NS', 'ASAHIINDIA.NS', 'GRANULES.NS', 'NAVA.NS', 'GSFC.NS', 'ALOKINDS.NS', 'CENTRALBK.NS', 'BANCOINDIA.NS', 
              'HONAUT.NS', 'EASEMYTRIP.NS', 'THOMASCOOK.NS', 'NELCO.NS', 'CREDITACC.NS', 'AIAENG.NS', 'DELTACORP.NS', 'TATACOMM.NS', 'GESHIP.NS', 'JUBLPHARMA.NS', 'EIDPARRY.NS', 'ZENTEC.NS', 'GPPL.NS', 'BORORENEW.NS', 'SOMATEX.NS', 'JWL.NS', 'PRICOLLTD.NS', 'GRWRHITECH.NS', 'ATUL.NS', 'JYOTISTRUC.NS', 'EMAMILTD.NS', 'GLOBECIVIL.NS', 'INSECTICID.NS', 'SKYGOLD.NS', 'LLOYDSME.NS', 'EMCURE.NS', 'DYNAMATECH.NS', 'WHIRLPOOL.NS', 'IKS.NS', 'LIQUIDADD.NS', 'RAIN.NS', 'INDIAMART.NS', 'ENDURANCE.NS', 'CUB.NS', 'HUBTOWN.NS', 'JKPAPER.NS', 'AKUMS.NS', 'SCODATUBES.NS', 'JKIL.NS', 'FLUOROCHEM.NS', 'JUBLFOOD.NS', 'UBL.NS', 'FAZE3Q.NS', 'JINDALSAW.NS', 'SBFC.NS', 'SCI.NS', 'BERGEPAINT.NS', 'JYOTHYLAB.NS', 'BIKAJI.NS', 'AFFLE.NS', 'BALRAMCHIN.NS', 'GRMOVER.NS', 'JSL.NS', 'LMW.NS', 'EPACK.NS', 'DHANI.NS', 'GNFC.NS', 'KIRLPNU.NS', 'TORNTPOWER.NS', 'CSBBANK.NS', 'BANKBEES.NS', 'GMBREW.NS', 'POONAWALLA.NS', 'AZAD.NS', 'WELSPUNLIV.NS', 'FEDFINA.NS', 'NIVABUPA.NS', 'AEGISLOG.NS', 'MOL.NS', 'TI.NS', 'ITIETF.NS', 'CHALET.NS', 'TBZ.NS', 'PARKHOTELS.NS', 'PHARMABEES.NS', 'AVANTIFEED.NS', 'RAINBOW.NS', 'SCHAEFFLER.NS', 'ORIENTHOT.NS', 'AVL.NS', 'SUPRAJIT.NS', 'IPL.NS', 'IOB.NS', 'NEWGEN.NS', 'FINEORG.NS', 'IT.NS', 'NETWORK18.NS', 'JISLJALEQS.NS', 'TECH.NS', 'SBIETFIT.NS', 'SENORES.NS', 'LIQUIDIETF.NS', 'AXISTECETF.NS', 'MARKSANS.NS', 'BFUTILITIE.NS', 'CHOLAHLDNG.NS', 'GPTHEALTH.NS', 'SILVERIETF.NS', 'GRINDWELL.NS', 'KSCL.NS', 'ACE.NS', 'RAMKY.NS', 'SIRCA.NS', 'DCXINDIA.NS', 'HERITGFOOD.NS', 'CAPLIPOINT.NS', '63MOONS.NS', 'MODEFENCE.NS', 'THELEELA.NS', 'INOXINDIA.NS', 'INTELLECT.NS', 'CRISIL.NS', 'KTKBANK.NS', 'DENTA.NS', 'MGL.NS', 'NAVINFLUOR.NS', 'PARSVNATH.NS', 'AKZOINDIA.NS', 'BATAINDIA.NS', 'APLLTD.NS', 'KSB.NS', 'MASFIN.NS', 'THYROCARE.NS', 'SPORTKING.NS', 'SIMPLEXINF.NS', 'VIJAYA.NS', 'NLCINDIA.NS', 'INFIBEAM.NS', 'ASKAUTOLTD.NS', 'KOKUYOCMLN.NS', 'WABAG.NS', 'GLAXO.NS', 'DIGITIDE.NS', 'AAVAS.NS', 'GPIL.NS', 'JUNIORBEES.NS', 'SETFNIF50.NS', 'REPCOHOME.NS', 'ANURAS.NS', 'DOLLAR.NS', 'BBL.NS', 'TARC.NS', 'ALIVUS.NS', 'BLUEDART.NS', 'ECLERX.NS', 'TIPSMUSIC.NS', 'EVEREADY.NS', 'GODREJIND.NS', 'TDPOWERSYS.NS', 'GHCL.NS', 'SBC.NS', 'DEEPAKNTR.NS', 'DDEVPLSTIK.NS', 'KIRIINDUS.NS', 'HERANBA.NS', 'FUSION.NS', 'PTCIL.NS', 'JAGSNPHARM.NS', 'MEDIASSIST.NS', 'POLYMED.NS', 'HUHTAMAKI.NS', 'BBTC.NS', 'SKIPPER.NS', 'SEPC.NS', 'VENUSPIPES.NS', 'RPGLIFE.NS', 'PGIL.NS', 'PRIVISCL.NS', 'SUNTV.NS', 'JKTYRE.NS', 'PANACEABIO.NS', 'MAPMYINDIA.NS', 'MAHLOG.NS', 'PRAKASH.NS', 'MANYAVAR.NS', 'CARERATING.NS', 'INDIASHLTR.NS', 'VASCONEQ.NS', 'ASHOKA.NS', 'GOODLUCK.NS', 'ALEMBICLTD.NS', 'PSUBANKADD.NS', 'GLOBUSSPR.NS', 'HARIOMPIPE.NS', 'MAHSCOOTER.NS', 'ABBOTINDIA.NS', 'RAYMONDLSL.NS', 'MANINDS.NS', 'MOIL.NS', 'CONCORDBIO.NS', 'VGUARD.NS', 'SUDARSCHEM.NS', 'VERANDA.NS', 'FINCABLES.NS', 'ORCHPHARMA.NS', 'SEQUENT.NS', 'HCG.NS', 'SUNDARMFIN.NS', 'LIQUID1.NS', 'SDBL.NS', 'SHYAMMETL.NS', 'KRISHANA.NS', 'ASTRAMICRO.NS', 'DODLA.NS', 'E2E.NS', 'SHOPERSTOP.NS', 'GTLINFRA.NS', 'ATHERENERG.NS', 'PFIZER.NS', 'PNGJL.NS', 'INFRAIETF.NS', 'DEN.NS', 'ACI.NS', 'CLEAN.NS', 'INDIAGLYCO.NS', 'SURYAROSNI.NS', 'ELGIEQUIP.NS', 'SJS.NS', 'JINDWORLD.NS', 'REFEX.NS', 'THERMAX.NS', 'CASHIETF.NS', 'FDC.NS', 'KALPATARU.NS', 'RPEL.NS', 'PSUBNKBEES.NS', 'MSTCLTD.NS', 'LANDMARK.NS', 'HDFCGOLD.NS', 'MANGCHEFER.NS', 'CMSINFO.NS', 'TRIVENI.NS', 'RITES.NS', 'VENTIVE.NS', 'SALASAR.NS', 'NIFTYIETF.NS', 'BAJAJHIND.NS', 'CEIGALL.NS', 'GODIGIT.NS', 'AURIONPRO.NS', 'RRKABEL.NS', 'AFCONS.NS', 'JAMNAAUTO.NS', 'ABREL.NS', 'MICEL.NS', 'CONTROLPR.NS', 'CCL.NS', 'DALBHARAT.NS', 'PSB.NS', 'GOLDIETF.NS', 'TIRUMALCHM.NS', 'BALAMINES.NS', 'AVALON.NS', 'PREMEXPLN.NS', 'GSPL.NS', 'BESTAGRO.NS', 'OPTIEMUS.NS', 'WEL.NS', 'HAPPSTMNDS.NS', 'SOLARA.NS', 'PGHH.NS', 'MOREPENLAB.NS', 'INDOSTAR.NS', 'AVANTEL.NS', 'LIQUIDETF.NS', 'TATAGOLD.NS', 'TRIDENT.NS', 'GENESYS.NS', 'NIF100BEES.NS', '3MINDIA.NS', 'CIGNITITEC.NS', 'SFL.NS', 'VSTIND.NS', 'NDRAUTO.NS', 'DYCL.NS', 'LATENTVIEW.NS', 'MAXESTATES.NS', 'EMUDHRA.NS', 'HITECH.NS', 'PFS.NS', 'PNBGILTS.NS', 'MEDPLUS.NS', 'TEAMLEASE.NS', 'FINPIPE.NS', 'EMSLIMITED.NS', 'IOLCP.NS', 'KPIL.NS', 'GOKEX.NS', 'JAIBALAJI.NS', 'GREENPOWER.NS', 'FLAIR.NS', 'SARDAEN.NS', 'NAVKARCORP.NS', 'BANKIETF.NS', 'SHRIPISTON.NS', 'INTLCONV.NS', 'HINDOILEXP.NS', 'KRBL.NS', 'WESTLIFE.NS', 'RELAXO.NS', 'SPARC.NS', 'ORIENTCEM.NS', 'MOSCHIP.NS', 'MPSLTD.NS', 'BANSALWIRE.NS', 'DCW.NS', 'LOWVOLIETF.NS', 'KNRCON.NS', 'UNICHEMLAB.NS', 'DEEPINDS.NS', 'RAMASTEEL.NS', 'BANKNIFTY1.NS', 'PRECWIRE.NS', 'SKFINDIA.NS', 'ROUTE.NS', 'AEROENTER.NS', 'VRLLOG.NS', 'JBMA.NS', 'AJAXENGG.NS', 'ZAGGLE.NS', 'CAMPUS.NS', 'SHAREINDIA.NS', 'DLINKINDIA.NS', 'IDEAFORGE.NS', 'PNCINFRA.NS', 'FIEMIND.NS', 'EMIL.NS', 'UNIVASTU.NS', 'METROBRAND.NS', 'HDFCNIFBAN.NS', 'FMGOETZE.NS', 'DAMCAPITAL.NS', 'ALKYLAMINE.NS', 'ABSLBANETF.NS', 'MID150BEES.NS', 'GOKULAGRO.NS', 'BECTORFOOD.NS', 'JTLIND.NS', 'TASTYBITE.NS', 'MAHSEAMLES.NS', 'SAKSOFT.NS', 'TVSSCS.NS', 'PROZONER.NS', 'SHALBY.NS', 'TNPETRO.NS', 'JCHAC.NS', 'SMALLCAP.NS', 'RBA.NS', 'KITEX.NS', 'FMCGIETF.NS', 'REDTAPE.NS', 'MADRASFERT.NS', 'ERIS.NS', 'HONASA.NS', 'MAYURUNIQ.NS', 'PRUDENT.NS', 'COMMOIETF.NS', 'RENUKA.NS', 'SANDUMA.NS', 'ANUP.NS', 'CARBORUNIV.NS', 'MINDACORP.NS', 'TATSILV.NS', 'HATHWAY.NS', 'QUESS.NS', 'BASF.NS', 'UDS.NS', 'HIKAL.NS', 'LXCHEM.NS', 'EIHOTEL.NS', 'AWFIS.NS', 'AGI.NS', 'NOCIL.NS', 'MIRCELECTR.NS', 'GUJTHEM.NS', 'INDOBORAX.NS', 'TATAINVEST.NS', 'VMART.NS', 'ZFCVINDIA.NS', 'EPIGRAL.NS', 'GUJGASLTD.NS', 'SUNFLAG.NS', 'HGINFRA.NS', 'GANECOS.NS', 'SILVERADD.NS', 'ARIES.NS', 'SANGHVIMOV.NS', 'INDRAMEDCO.NS', 'ARVIND.NS', 'PRINCEPIPE.NS', 'SHAILY.NS', 'SILVER.NS', 'POLYPLEX.NS', 'ECOSMOBLTY.NS', 'DECCANCE.NS', 'MON100.NS', 'VTL.NS', 'GOLDIAM.NS', 'SHRIRAMPPS.NS', 'EUREKAFORB.NS', 'FINOPB.NS', 'BAJAJCON.NS', 'CYIENTDLM.NS', 'CPSEETF.NS', 'CAPACITE.NS', 'TAJGVK.NS', 'RADHIKAJWE.NS', 'GMMPFAUDLR.NS', 'PATELENG.NS', 'LIQUID.NS', 'LINDEINDIA.NS', 'WINDMACHIN.NS', 'UNIECOM.NS', 'GODREJAGRO.NS', 'INDIACEM.NS', 'SPIC.NS', 'RATNAMANI.NS', 'BLUSPRING.NS', 'EXICOM.NS', 'SUBROS.NS', 'SETFGOLD.NS', 'SARVESHWAR.NS', 'SUNDARMHLD.NS', 'COMSYN.NS', 'ARVINDFASN.NS', 'MOTISONS.NS', 'MUFIN.NS', 'STERTOOLS.NS', 'MONARCH.NS', 
              'HLEGLAS.NS', 'VINATIORGA.NS', 'APOLLOPIPE.NS', 'GOLDETFADD.NS', 'JNKINDIA.NS', 'ZYDUSWELL.NS', 'ICIL.NS', 'IXIGO.NS', 'MANINFRA.NS', 'RSYSTEMS.NS', 'STYRENIX.NS', 'IIFLCAPS.NS', 'SYMPHONY.NS', 'NIITMTS.NS', 'SURAJEST.NS', 'ROTO.NS', 'SHANKARA.NS', 'SSWL.NS', 'SINDHUTRAD.NS', 'UTKARSHBNK.NS', 'SUNDRMFAST.NS', 'INGERRAND.NS', 'CENTURYPLY.NS', 'ARVSMART.NS', 'SUNTECK.NS', 'PIXTRANS.NS', 'HDFCNIFTY.NS', 'IFBIND.NS', 'VERTOZ.NS', 'SANATHAN.NS', 'MTARTECH.NS', 'BAJAJINDEF.NS', 'ITI.NS', 'GEOJITFSL.NS', 'DBOL.NS', 'NSLNISP.NS', 'RHIM.NS', 'IMFA.NS', 'GOCOLORS.NS', 'VAIBHAVGBL.NS', 'HDFCPSUBK.NS', 'SUVIDHAA.NS', 'SILVER1.NS', 'GIPCL.NS', 'EXCELINDUS.NS', 'RIIL.NS', 'ALPHA.NS', 'LGHL.NS', 'SAREGAMA.NS', 'SILVERETF.NS', 'MSPL.NS', 'MONIFTY500.NS', 'HEMIPROP.NS', 'COSMOFIRST.NS', 'SANSERA.NS', 'LTGILTBEES.NS', 'RATNAVEER.NS', 'AGARWALEYE.NS', 'BAJAJELEC.NS', 'SHREEPUSHK.NS', 'GULFOILLUB.NS', 'JTEKTINDIA.NS', 'ADVENZYMES.NS', 'WELENT.NS', 'AVADHSUGAR.NS', 'STEELXIND.NS', 'GICHSGFIN.NS', 'GATEWAY.NS', 'AXISNIFTY.NS', 'VIMTALABS.NS', 'SURYODAY.NS', 'ZOTA.NS', 'NPST.NS', 'JSFB.NS', 'WSTCSTPAPR.NS', 'HIMATSEIDE.NS', 'CARYSIL.NS', 'KOPRAN.NS', 'INNOVACAP.NS', 'QUICKHEAL.NS', 'BOROLTD.NS', 'GARFIBRES.NS', 'VSTTILLERS.NS', 'KMEW.NS', 'STYLAMIND.NS', 'THEJO.NS', 'PITTIENG.NS', 'LIQUIDPLUS.NS', 'SULA.NS', 'SIS.NS', 'ASALCBR.NS', 'CANTABIL.NS', 'PGHL.NS', 'KELLTONTEC.NS', 'AWHCL.NS', 'KSL.NS', 'BHARATWIRE.NS', 'SMSPHARMA.NS', 'LUMAXTECH.NS', 'LOKESHMACH.NS', 'CELLO.NS', 'NIACL.NS', 'HIRECT.NS', 'ZEEMEDIA.NS', 'SANDHAR.NS', 'DBL.NS', 'DPABHUSHAN.NS', 'BIRLACORPN.NS', 'INDSWFTLTD.NS', 'KOLTEPATIL.NS', 'TCI.NS', 'UGROCAP.NS', 'PUNJABCHEM.NS', 'EKC.NS', 'ETHOSLTD.NS', 'CENTUM.NS', 'FILATFASH.NS', 'GILT5YBEES.NS', 'ARISINFRA.NS', 'DSSL.NS', 'VESUVIUS.NS', 'MOREALTY.NS', 'AUTOBEES.NS', 'PRSMJOHNSN.NS', 'PARACABLES.NS', 'TARACHAND.NS', 'VISHNU.NS', 'NSIL.NS', 'SBCL.NS', 'FILATEX.NS', 'GALLANTT.NS', 'EPL.NS', 'LUXIND.NS', 'PSUBNKIETF.NS', 'V2RETAIL.NS', 'ESILVER.NS', 'UNIVCABLES.NS', 'DHANBANK.NS', 'SALZERELEC.NS', 'DIAMONDYD.NS', 'BIRLAMONEY.NS', 'KHADIM.NS', 'TMB.NS', 'PRABHA.NS', 'SGIL.NS', 'SIYSIL.NS', 'CANFINHOME.NS', 'IONEXCHANG.NS', 'NRBBEARING.NS', 'KABRAEXTRU.NS', 'BHAGCHEM.NS', 'KRSNAA.NS', 'PENINLAND.NS', 'SRD.NS', 'MONTECARLO.NS', 'VENKEYS.NS', 'STYLEBAAZA.NS', 'KDDL.NS', 'ACLGATI.NS', 'CIEINDIA.NS', 'HNGSNGBEES.NS', 'ARIHANTSUP.NS', 'HONDAPOWER.NS', 'KALAMANDIR.NS', 'SHIVALIK.NS', 'PSUBANK.NS', 'VADILALIND.NS', 'ISGEC.NS', 'MOM30IETF.NS', 'KAMDHENU.NS', 'SIGNPOST.NS', 'MMTC.NS', 'ORIENTTECH.NS', 'REMSONSIND.NS', 'HEIDELBERG.NS', 'NEXT50IETF.NS', 'PURVA.NS', 'CLSEL.NS', 'JASH.NS', 'KANSAINER.NS', 'AXISILVER.NS', 'SBISILVER.NS', 'MAHKTECH.NS', 'PRECAM.NS', 'MAFANG.NS', 'PDSL.NS', 'ZUARI.NS', 'GRINFRA.NS', 'MAITHANALL.NS', 'GROWWDEFNC.NS', 'THANGAMAYL.NS', 'BHARATRAS.NS', 'OMAXAUTO.NS', 'KRONOX.NS', 'ANDHRSUGAR.NS', 'BODALCHEM.NS', 'ARMANFIN.NS', 'SRHHYPOLTD.NS', 'SHARDAMOTR.NS', 'TATVA.NS', 'WHEELS.NS', 'JUNIPER.NS', 'BLISSGVS.NS', 'STOVEKRAFT.NS', 'IMAGICAA.NS', 'GREENPANEL.NS', 'NIITLTD.NS', 'NEOGEN.NS', 'HARSHA.NS', 'SCILAL.NS', 'SHK.NS', 'SELAN.NS', 'INDIGOPNTS.NS', 'MUTHOOTMF.NS', 'HNDFDS.NS', 'AJMERA.NS', 'ANUHPHR.NS', 'TNPL.NS', 'DALMIASUG.NS', 'LGBBROSLTD.NS', 'SANGAMIND.NS', 'FOSECOIND.NS', 'OSWALAGRO.NS', 'GAEL.NS', 'MAHLIFE.NS', 'MAZDA.NS', 'GALAPREC.NS', 'HPAL.NS', 'MOCAPITAL.NS', 'LIQUIDBETF.NS', 'CONFIPET.NS', 'CIFL.NS', 'AVONMORE.NS', 'ARTEMISMED.NS', 'SANOFI.NS', 'POKARNA.NS', 'RPTECH.NS', 'RGL.NS', 'SAPPHIRE.NS', 'ASAL.NS', 'ORISSAMINE.NS', 'URJA.NS', 'SRM.NS', 'AAATECH.NS', 'DEEDEV.NS', 'METALIETF.NS', 'ICRA.NS', 'MVGJL.NS', 'OMAXE.NS', 'ICEMAKE.NS', 'JINDRILL.NS', 'ATULAUTO.NS', 'MOSMALL250.NS', 'KROSS.NS', 'PAKKA.NS', 'TALBROAUTO.NS', 'BORANA.NS', 'GREENPLY.NS', 'VINDHYATEL.NS', 'DBCORP.NS', 'KCP.NS', 'GOPAL.NS', 'MOMENTUM50.NS', 'ORIENTELEC.NS', 'SHREDIGCEM.NS', 'NCLIND.NS', 'GOCLCORP.NS', 'NITCO.NS', 'STEELCAS.NS', 'GOLDCASE.NS', 'NAVNETEDUL.NS', 'ENTERO.NS', 'NOVAAGRI.NS', 'JAGRAN.NS', 'BFINVEST.NS', 'SANOFICONR.NS', 'PICCADIL.NS', 'MUTHOOTCAP.NS', 'BALMLAWRIE.NS', 'TIIL.NS', 'EXPLEOSOL.NS', 'TEXINFRA.NS', 'TOP10ADD.NS', 'MEDICAMEQ.NS', 'KECL.NS', 'RUSTOMJEE.NS', 'UDAICEMENT.NS', 'HAPPYFORGE.NS', 'FCL.NS', 'MANALIPETC.NS', 'MOM100.NS', 'TTKPRESTIG.NS', 'TCPLPACK.NS', 'HEUBACHIND.NS', 'RITCO.NS', 'DOLATALGO.NS', 'ROLEXRINGS.NS', 'MIDCAPETF.NS', 'WINDLAS.NS', 'JAICORPLTD.NS', 'SAFARI.NS', 'DCMSRIND.NS', 'BLKASHYAP.NS', 'INDOAMIN.NS', 'VASWANI.NS', 'LUMAXIND.NS', 'ATL.NS', 'INDIANHUME.NS', 'PLATIND.NS', 'XPROINDIA.NS', 'ADFFOODS.NS', 'SANSTAR.NS', 'IFGLEXPOR.NS', 'INDOTECH.NS', 'WEALTH.NS', 'ASTEC.NS', 'GALAXYSURF.NS', 'AVROIND.NS', 'GNA.NS', 'BCLIND.NS', 'ICICIB22.NS', 'NILASPACES.NS', 'BARBEQUE.NS', 'BEPL.NS', 'RPSGVENT.NS', 'GUFICBIO.NS', 'VEEDOL.NS', 'WCIL.NS', 'DHAMPURSUG.NS', 'PDMJEPAPER.NS', 'RML.NS', 'NITINSPIN.NS', 'COFFEEDAY.NS', 'VIKASLIFE.NS', 'BHARATGEAR.NS', 'MCLEODRUSS.NS', 'MONQ50.NS', 'MGEL.NS', 'ADSL.NS', 'SPAL.NS', 'CEWATER.NS', 'SILVERCASE.NS', 'GUJAPOLLO.NS', 'FINIETF.NS', 'MANGLMCEM.NS', 'APCOTEXIND.NS', 'INDOCO.NS', 'MUNJALSHOW.NS', 'ASMS.NS', 'NIBE.NS', 'GPTINFRA.NS', 'JLHL.NS', 'DREDGECORP.NS', 'KKCL.NS', 'GOLDETF.NS', 'INDNIPPON.NS', 'RUBYMILLS.NS', 'SMCGLOBAL.NS', 'ROHLTD.NS', 'LAOPALA.NS', 'SANGHIIND.NS', 'SATIN.NS', 'MMFL.NS', 'JUBLCPL.NS', 'CARRARO.NS', 'EVINDIA.NS', 'AXITA.NS', 'TRACXN.NS', 'MALLCOM.NS', 'ITETF.NS', 'CONSUMBEES.NS', 'AUTOIETF.NS', 'PVTBANIETF.NS', 'GUJALKALI.NS', 'MASPTOP50.NS', 'GOLD1.NS', 'AMBICAAGAR.NS', 'ASHIANA.NS', 'LINCOLN.NS', 'BALAJEE.NS', 'HDFCLIQUID.NS', 'NUCLEUS.NS', 'MOLDTECH.NS', 'KIOCL.NS', 'EIMCOELECO.NS', 'INNOVANA.NS', 'RESPONIND.NS', 'WONDERLA.NS', 'AMBIKCO.NS', 'SILVRETF.NS', '20MICRONS.NS', 'LIKHITHA.NS', 'RAJRATAN.NS', 'THEMISMED.NS', 'SAMPANN.NS', 'ITDC.NS', 'ROSSTECH.NS', 'DIVGIITTS.NS', 'ESAFSFB.NS', 'DYNPRO.NS', 'MUFTI.NS', 'SUNDROP.NS', 'BALAXI.NS', 'SNOWMAN.NS', 'AGARIND.NS', 'RICOAUTO.NS', 'BIL.NS', 'SAGCEM.NS', 'BIRLANU.NS', 'ABSLNN50ET.NS', 'MIDCAPIETF.NS', 'GULPOLY.NS', 'ANDHRAPAP.NS', 'ASIANENE.NS', 'PVSL.NS', 'MOLDTKPAC.NS', 'PPL.NS', 'EBBETF0433.NS', 'MUNJALAU.NS', 'MID150CASE.NS', 'KINGFA.NS', 'ORIENTPPR.NS', 'NDL.NS', 'ZUARIIND.NS', 'SETFNN50.NS', 'EBBETF0430.NS', 'HATSUN.NS', 'ESTER.NS', 'HESTERBIO.NS', 'VIRINCHI.NS', 'RACLGEAR.NS', 'GROWWGOLD.NS', 'GHCLTEXTIL.NS', 'RADIANTCMS.NS', 'MBLINFRA.NS', 'GROWWSLVR.NS', 'PILANIINVS.NS', 'SASKEN.NS', 'HINDWAREAP.NS', 'SOMANYCERA.NS', 'YATRA.NS', 'RANEHOLDIN.NS', 'RUSHIL.NS', 'GANDHAR.NS', 'GOACARBON.NS', 'NILAINFRA.NS', 'GROWWRAIL.NS', 'UNIPARTS.NS', 'HEALTHIETF.NS', 'AFSL.NS', 'SUNCLAY.NS', 'MHRIL.NS', 'MANAKCOAT.NS', 'ESABINDIA.NS', 'KAMATHOTEL.NS', 'GANESHBE.NS', 'UMIYA-MRO.NS', 'NARMADA.NS', 'SHANTIGEAR.NS', 'SADBHAV.NS', 'ALANKIT.NS', 'VIPCLOTHNG.NS', 'SUPREME.NS', 'MANCREDIT.NS', 'DIAMINESQ.NS', 'APOLSINHOT.NS', 'SILVERTUC.NS', 'UTTAMSUGAR.NS', 'SUMMITSEC.NS', 'DWARKESH.NS', 'ACCELYA.NS', 'GROWWLIQID.NS', 'AXISGOLD.NS', 'IRMENERGY.NS', 'AONELIQUID.NS', 'LORDSCHLO.NS', 'RPPINFRA.NS', 'KSOLVES.NS', 'JINDALPOLY.NS', '5PAISA.NS', 'TVSSRICHAK.NS', 'AURUM.NS', 'IKIO.NS', 'BIGBLOC.NS', 'MUKANDLTD.NS', 'SURAKSHA.NS', 'NAHARSPING.NS', 'RAMRAT.NS', 'CROWN.NS', 'STARCEMENT.NS', 'NIFTYETF.NS', 'OSWALGREEN.NS', 'GROWWEV.NS', 'GTPL.NS', 'NGLFINE.NS', 'ALPL30IETF.NS', 'RBZJEWEL.NS', 'AARTECH.NS', 'SPLPETRO.NS', 'OMINFRAL.NS', 'BHAGYANGR.NS', 'MEDICO.NS', 'BSLNIFTY.NS', 'MIDSMALL.NS', 'LIQUIDSHRI.NS', 'SOTL.NS', 'YASHO.NS', 'CENTEXT.NS', 'ARIHANTCAP.NS', 'STANLEY.NS', 'MMP.NS', 'KUANTUM.NS', 'BAIDFIN.NS', 'HMVL.NS', 'ALLDIGI.NS', 'TEXMOPIPES.NS', 'RAMCOIND.NS', 'TCIEXP.NS', 'SHALPAINTS.NS', 'IZMO.NS', 'CHEMPLASTS.NS', 'GOLDSHARE.NS', 'MAXIND.NS', 'PRIMESECU.NS', 'HDFCMID150.NS', 'VSTL.NS', 'ALPHAETF.NS', 'NIFTY1.NS', 'HGS.NS', 'NV20IETF.NS', 'NELCAST.NS', 'GSS.NS', 'DCM.NS', 'VAISHALI.NS', 'UCAL.NS', 'MANBA.NS', 'RISHABH.NS', 'SETFNIFBK.NS', 'NEXT50.NS', 'AONETOTAL.NS', 'SESHAPAPER.NS', 'JAYSREETEA.NS', 'OILIETF.NS', 'CAPITALSFB.NS', 'MAGADSUGAR.NS', 'VIDHIING.NS', 'VIKASECO.NS', 'XCHANGING.NS', 'NIPPOBATRY.NS', 'AUTOAXLES.NS', 'RKSWAMY.NS', 'SONAMLTD.NS', 'FAIRCHEMOR.NS', 'TOLINS.NS', 'TREL.NS', '3IINFOLTD.NS', 'EMKAY.NS', 'IGPL.NS', 'BBTCL.NS', 'KICL.NS', 'MUKKA.NS', 'SCPL.NS', 'ONMOBILE.NS', 'MODISONLTD.NS', 'ORIENTCER.NS', 'VRAJ.NS', 'GRPLTD.NS', 'LOWVOL1.NS', 'ALICON.NS', 'SPENCERS.NS', 'YUKEN.NS', 'SATIA.NS', 'PLASTIBLEN.NS', 'MAHEPC.NS', 'CCCL.NS', 'TARSONS.NS', 'DBEIL.NS', 'SHARIABEES.NS', 'BOROSCI.NS', 'UFO.NS', 'AUTOIND.NS', 'TRF.NS', 'RADIOCITY.NS', 'TRANSWORLD.NS', 'MAWANASUG.NS', 'SUTLEJTEX.NS', 'KCPSUGIND.NS', 'MOVALUE.NS', 'TVTODAY.NS', 'KOTHARIPET.NS', 'GSLSU.NS', 'GFLLIMITED.NS', 'ALMONDZ.NS', 'DPWIRES.NS', 'FISCHER.NS', 'AARON.NS', 'TPLPLASTEH.NS', 'CENTENKA.NS', 'BPL.NS', 'AVTNPL.NS', 'MAGNUM.NS', 'VSSL.NS', 'APEX.NS', 'NURECA.NS', 'SAURASHCEM.NS', 'MOMOMENTUM.NS', 'ZIMLAB.NS', 'SEAMECLTD.NS', 'HDFCMOMENT.NS', 'BANARISUG.NS', 'AMRUTANJAN.NS', 'HMAAGRO.NS', 'ADOR.NS', 'UFLEX.NS', 'PASUPTAC.NS', 'UGARSUGAR.NS', 'HLVLTD.NS', 'RUCHIRA.NS', 'SUPERSPIN.NS', 'CHEVIOT.NS', 'MHLXMIRU.NS', 'GENUSPAPER.NS', 'TARMAT.NS', 'UNITEDPOLY.NS', 'DENORA.NS', 'SUMIT.NS', 'ABSLLIQUID.NS', 'BFSI.NS', 'DAVANGERE.NS', 'RUBFILA.NS', 'GTL.NS', 'MENONBE.NS', 'CLEDUCATE.NS', 'BSLGOLDETF.NS', 'AVG.NS', 'VETO.NS', 'STARTECK.NS', 'HDFCNIFIT.NS', 'PANAMAPET.NS', 'INVENTURE.NS', 'SWELECTES.NS', 'ALBERTDAVD.NS', 'BUTTERFLY.NS', 'BHAGERIA.NS', 'FCSSOFT.NS', 'GEEKAYWIRE.NS', 'DONEAR.NS', 'TTKHLTCARE.NS', 'AGRITECH.NS', 'EIHAHOTELS.NS', 'CREATIVE.NS', 'JITFINFRA.NS', 'SOFTTECH.NS', 'GLOBALVECT.NS', 'INDOWIND.NS', 'OAL.NS', 'PTL.NS', 'SENSEXIETF.NS', 'SCHAND.NS', 'MASTERTR.NS', 'MIRZAINT.NS', 'ITETFADD.NS', 'SASTASUNDR.NS', 'BLAL.NS', 'REPRO.NS', 'METAL.NS', 'MIDCAP.NS', 'JAYAGROGN.NS', 'GULFPETRO.NS', 'PRITIKAUTO.NS', 'NILKAMAL.NS', 'GLOBE.NS', 'SAKHTISUG.NS', 'RPPL.NS', 'INFRABEES.NS', 'PVP.NS', 'AFIL.NS', 'INTERNET.NS', 'EMAMIPAP.NS', 'BHANDARI.NS', 'MIDQ50ADD.NS', 'ABAN.NS', 'PAVNAIND.NS', 'VLSFINANCE.NS', 'AFFORDABLE.NS', 'LIQUIDSBI.NS', 'SUKHJITS.NS', 'LICMFGOLD.NS', 'SALONA.NS', 'TOP100CASE.NS', 'MATRIMONY.NS', 'UNIDT.NS', 'CORDSCABLE.NS', 'SHREERAMA.NS', 'SETF10GILT.NS', 'AKSHOPTFBR.NS', 'INTENTECH.NS', 'AGROPHOS.NS', 'ORIENTBELL.NS', 'ADVANIHOTR.NS', 'TNIDETF.NS', 'CHEMCON.NS', 'ENIL.NS', 'PREMIERPOL.NS', 'AARVI.NS', 'MOKSH.NS', 'SINCLAIR.NS', 'CHEMFAB.NS', 'HDFCNEXT50.NS', 'STEELCITY.NS', 'LOVABLE.NS', 'AONENIFTY.NS', 'CONSUMIETF.NS', 'JAYBARMARU.NS', 'BASML.NS', 'RKEC.NS', 'KOHINOOR.NS', 'GREENLAM.NS', 'AJOONI.NS', 'PYRAMID.NS', 'DEVIT.NS', 'HEALTHY.NS', 'AKG.NS', 'NIF100IETF.NS', 'GEECEE.NS', 'ESSENTIA.NS', 'LYKALABS.NS', 'EMULTIMQ.NS', 'EVIETF.NS', 'GANDHITUBE.NS', 'MCL.NS', 'INFOBEAN.NS', 'AYMSYNTEX.NS', 'NV20BEES.NS', 'DJML.NS', 'ELECTHERM.NS', 'SURANASOL.NS', 'GOKUL.NS', 'RSWM.NS', 'AKSHARCHEM.NS', 'AIRAN.NS', 'NINSYS.NS', 'LOYALTEX.NS', 'IMPAL.NS', 'EXCEL.NS', 'DPSCLTD.NS', 'JPOLYINVST.NS', 'BANSWRAS.NS', 'DIVOPPBEES.NS', 'AROGRANITE.NS', 'MNC.NS', 'UNITEDTEA.NS', 'QUAL30IETF.NS', 'GROWWMOM50.NS', 'JHS.NS', 'LATTEYS.NS', 'RANASUG.NS', 'SREEL.NS', 'BSHSL.NS', 'PRUDMOULI.NS', 'SMLT.NS', 'DVL.NS', 'SSDL.NS', 'FIBERWEB.NS', 'ISFT.NS', 'MOHEALTH.NS', 'SPECTRUM.NS', 'BALPHARMA.NS', 'MANGALAM.NS', 'WORTH.NS', 'SHIVAMAUTO.NS', 'SIKKO.NS', 'HITECHCORP.NS', 'ALPA.NS', 'HISARMETAL.NS', 'PLAZACABLE.NS', 'PONNIERODE.NS', 'GROWWNET.NS', 'TREJHARA.NS', 'RVHL.NS', 'HDFCSENSEX.NS', 'GLOSTERLTD.NS', 'IVP.NS', 'BANKETFADD.NS', 'SECURKLOUD.NS', 'DCI.NS', 'VHL.NS', 'UTINEXT50.NS', 'PATINTLOG.NS', 'DOLPHIN.NS', 'PRIMO.NS', 'MOMENTUM.NS', 'TOUCHWOOD.NS', 'VALIANTLAB.NS', 'MITTAL.NS', 'MIDSELIETF.NS', 'STARPAPER.NS', 'GANGAFORGE.NS', 'INDIANCARD.NS', 'VARDMNPOLY.NS', 'DIGISPICE.NS', 'LFIC.NS', 'AMJLAND.NS', 'PRAENG.NS', 'NECCLTD.NS', 'LANCORHOL.NS', 'MARALOVER.NS', 'REPL.NS', 'TIPSFILMS.NS', 'DTIL.NS', 'SAGARDEEP.NS', 'LOTUSEYE.NS', 'DIGIDRIVE.NS', 'CELEBRITY.NS', 'DUCON.NS', 'SURANAT&P.NS', 'MALUPAPER.NS', 'TRIGYN.NS', 'HEXATRADEX.NS', 'PODDARMENT.NS', 'VISHWARAJ.NS', 'CCHHL.NS', 'KANORICHEM.NS', 'MEGASTAR.NS', 'BCONCEPTS.NS', 'UYFINCORP.NS', 'XELPMOC.NS', 'UTIBANKETF.NS', 'MURUDCERA.NS', 'BANG.NS', 'AMNPLST.NS', 'USK.NS', 'SMSLIFE.NS', 'DCMNVL.NS', 'MAHESHWARI.NS', 'SPECIALITY.NS', 'QGOLDHALF.NS', 'GILLANDERS.NS', 'KMSUGAR.NS', 'CSLFINANCE.NS', 'UTINIFTETF.NS', 'SENSEXETF.NS', 'ANMOL.NS', 'GROWWN200.NS', 'BSE500IETF.NS', 'NITIRAJ.NS', 'HINDCON.NS', 'NIFTY50ADD.NS', 'ROSSELLIND.NS', 'RELCHEMQ.NS', 'IVC.NS', 'THEINVEST.NS', 'COASTCORP.NS', 'HINDCOMPOS.NS', 'DHUNINV.NS', 'CREST.NS', 'KRITIKA.NS', 'MULTICAP.NS', 'ACL.NS', 'BEDMUTHA.NS', 'NIFTYQLITY.NS', 'KAPSTON.NS', 'VIJIFIN.NS', 'LAGNAM.NS', 'SURAJLTD.NS', 'ARCHIDPLY.NS', 'LICNMID100.NS', 'BEARDSELL.NS', 'ATLANTAA.NS', 'SAHYADRI.NS', 'GOLDTECH.NS', 'MITCON.NS', 'SADBHIN.NS', 'BANKPSU.NS', 'EMMBI.NS', 'SABTNL.NS', 'HARRMALAYA.NS', 'ACCURACY.NS', 'NV20.NS', 'ABSLPSE.NS', 'KALYANIFRG.NS', 'MAHAPEXLTD.NS', 'SILVER360.NS', 'ALPHAGEO.NS', 'KOTARISUG.NS', 'STEL.NS', 'SANDESH.NS', 'RHL.NS', 'KRITINUT.NS', 'HDFCPVTBAN.NS', 'SALSTEEL.NS', 'BTML.NS', 'RACE.NS', 'HDFCGROWTH.NS', 'INDOUS.NS', 'SBIETFPB.NS', 'PRAKASHSTL.NS', 'AIROLAM.NS', 'HDFCQUAL.NS', 'INDBANK.NS', 'MAKEINDIA.NS', 'MANAKSIA.NS', 'MOM50.NS', 'KSHITIJPOL.NS', 'AXISHCETF.NS', 'SUNDRMBRAK.NS', 'ZENITHEXPO.NS', 'BLBLIMITED.NS', 'MUKTAARTS.NS', 'TIMESGTY.NS', 'SBINEQWETF.NS', 'KAUSHALYA.NS', 'BANKETF.NS', 'ATAM.NS', 'SURYALAXMI.NS', 'BAFNAPH.NS', 'AMDIND.NS', 'LIBAS.NS', 'HDFCBSE500.NS', 'INDTERRAIN.NS', 'BAGFILMS.NS', 'PEARLPOLY.NS', 'VARDHACRLC.NS', 'GOYALALUM.NS', 'SHRENIK.NS', 'UTISENSETF.NS', 'OILCOUNTUB.NS', 'SHAHALLOYS.NS', 'TOP15IETF.NS', 'BVCL.NS', 'NBIFIN.NS', 'SILINV.NS', 'RETAIL.NS', 'SHRADHA.NS', 'ESG.NS', 'NIFTY100EW.NS', 'NIFITETF.NS', 'CORALFINAC.NS', 'SELECTIPO.NS', 'SHIVATEX.NS', 'JISLDVREQS.NS', 'GVPTECH.NS', 'EQUAL50ADD.NS', 'SILLYMONKS.NS', 'VAL30IETF.NS', 'ALKALI.NS', 'NIRAJ.NS', '21STCENMGM.NS', 'IVZINGOLD.NS', 'SUPERHOUSE.NS', 'OSWALSEEDS.NS', 'PRITI.NS', 'AHLADA.NS', 'PVTBANKADD.NS', 'LAMBODHARA.NS', 'SAMBHAAV.NS', 'CINELINE.NS', 'RAJSREESUG.NS', 'JMA.NS', 'NIFTYBETF.NS', 'SUNDARAM.NS', 'SIL.NS', 'DICIND.NS', 'GKWLIMITED.NS', 'COMPUSOFT.NS', 'HDFCVALUE.NS', 'SMARTLINK.NS', 'NATCAPSUQ.NS', 'ANTGRAPHIC.NS', 'JAIPURKURT.NS', 'RELIABLE.NS', 'AXISVALUE.NS', 'XTGLOBAL.NS', 'ANIKINDS.NS', 'KAKATCEM.NS', 'CONSOFINVT.NS', 'HDFCNIF100.NS', 'WEWIN.NS', 'TGBHOTELS.NS', 'TAINWALCHM.NS', 'SBIBPB.NS', 'NDLVENTURE.NS', 'AHLEAST.NS', 'BOHRAIND.NS', 'GSEC10IETF.NS', 'MOMIDMTM.NS', 'MOGSEC.NS', 'BALKRISHNA.NS', 'WEIZMANIND.NS', 'NDGL.NS', 'HDFCLOWVOL.NS', 'ADROITINFO.NS', 'OBCL.NS', 'SBIETFQLTY.NS', 'ASTRON.NS', 'EQUAL200.NS', 'ZODIACLOTH.NS', 'RKDL.NS', 'PARASPETRO.NS', 'DBSTOCKBRO.NS', 'ORIENTLTD.NS', 'GROWWLOVOL.NS', 'VIVIDHA.NS', 'JETFREIGHT.NS', 'KOTHARIPRO.NS', 'SOMICONVEY.NS', 'ECAPINSURE.NS', 'MADHAV.NS', 'SNXT30BEES.NS', 'AXISBNKETF.NS', 'BANKBETF.NS', 'UNIENTER.NS', 'TOTAL.NS', 'HEADSUP.NS', 'HPIL.NS', 'MSCIINDIA.NS', 'SINTERCOM.NS', 'CPCAP.NS', 'MANAKALUCO.NS', 'KANANIIND.NS', 'EGOLD.NS', '3PLAND.NS', 'MID150.NS', 'REGENCERAM.NS', 'CONSUMER.NS', 'BANARBEADS.NS', 'LAXMICOT.NS', 'TOKYOPLAST.NS', 'SBIETFCON.NS', 'SANGINITA.NS', 'BBNPPGOLD.NS', 'JOCIL.NS', 'NETF.NS', 'MOTOGENFIN.NS', 'ASHOKAMET.NS', 'LPDC.NS', 'QNIFTY.NS', 'RADAAN.NS', 'AXISBPSETF.NS', 'UTISXN50.NS', 'SHIVAMILLS.NS', 'NAGREEKEXP.NS', 'HEALTHADD.NS', 'SENSEXADD.NS', 'MANOMAY.NS', 'MOTOUR.NS', 'DNAMEDIA.NS', 'GSEC5IETF.NS', 'ASPINWALL.NS', 'ELDEHSG.NS', 'GUJRAFFIA.NS', 'BANKA.NS', 'MOHITIND.NS', 'SVPGLOB.NS', 'PALASHSECU.NS', 'CYBERMEDIA.NS', 'WIPL.NS', 'NIFMID150.NS', 'NGIL.NS', 'AKASH.NS', 'LEXUS.NS', 'SRGHFL.NS']

# Configurable thresholds (KEEPING YOUR ORIGINAL VALUES)
strong_buy_threshold = 0.90
buy_threshold = 0.70
strong_sell_threshold = 0.60
sell_threshold = 0.50

# Suppress future warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def get_ist_time():
    return datetime.now(IST)

def load_trading_log():
    """Load trading log from Excel file"""
    if os.path.exists(LOG_FILE):
        try:
            df = pd.read_excel(LOG_FILE)
            # Convert Date column to string for consistent comparison
            if 'Date' in df.columns:
                df['Date'] = df['Date'].astype(str)
            return df
        except Exception as e:
            print(f"Error loading trading log: {e}")
            return pd.DataFrame()
    
    return pd.DataFrame(columns=[
        'Date', 'Ticker', 'Buy_Price', 'Current_Price', 'Sell_Price', 
        'PnL', 'Status', 'Buy_Prob', 'Sell_Signal_Time', 'Investment_Amount',
        'Profit_Amount', 'Return_Percentage'
    ])

def save_trading_log(df):
    """Save trading log to Excel file"""
    try:
        # Calculate profit amounts before saving
        if 'Investment_Amount' not in df.columns:
            df['Investment_Amount'] = 100000  # 1 lakh per stock
            
        # Calculate profit amount and return percentage
        df['Profit_Amount'] = df['PnL'] * (df['Investment_Amount'] / df['Buy_Price'])
        df['Return_Percentage'] = (df['PnL'] / df['Buy_Price']) * 100
        
        df.to_excel(LOG_FILE, index=False)
        print(f"✓ Log saved to {LOG_FILE}")
        return True
    except Exception as e:
        print(f"✗ Error saving trading log: {e}")
        return False

def fetch_latest_data(ticker, period="1y", interval="1d"):
    """Fetch stock data with enhanced error handling and retry logic
    (YOUR ORIGINAL FUNCTION - COMPLETELY UNCHANGED)"""
    max_retries = 1
    retry_delay = 1

    for attempt in range(max_retries):
        try:
            stock = yf.Ticker(ticker)
            df = stock.history(period=period, interval=interval)

            if df.empty:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                return pd.DataFrame()

            df = df.reset_index()
            
            # Ensure date column exists
            if 'Date' not in df.columns and df.index.name == 'Date':
                df['Date'] = df.index
            elif 'Date' not in df.columns:
                df['Date'] = pd.to_datetime(df.index)
            
            # Clean data
            df = df.dropna()
            
            #print(f"Data fetched for {ticker} - {len(df)} rows")
            return df
            
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
            else:
                return pd.DataFrame()

def fetch_live_price(ticker):
    """Fetch current live price for monitoring"""
    try:
        stock = yf.Ticker(ticker)
        data = stock.history(period='1d', interval='1m')
        if not data.empty:
            return data['Close'].iloc[-1]
    except Exception as e:
        print(f"Error fetching live price for {ticker}: {e}")
    return None

def analyze_stock(df):
    """YOUR ORIGINAL analyze_stock FUNCTION - COMPLETELY UNCHANGED"""
    df.columns = [col[0] if isinstance(col, tuple) else col for col in df.columns.values]
    df.columns = [str(col).strip().lower() for col in df.columns]

    # Handle date column
    date_col = None
    for col in df.columns:
         if col.lower() == 'date':
             date_col = col
             break
            
    if date_col is None:
         if df.index.name == 'Date' or isinstance(df.index, pd.DatetimeIndex):
             df['date'] = pd.to_datetime(df.index)
         else:
             df['date'] = pd.to_datetime(df.index)
    else:
         df['date'] = pd.to_datetime(df[date_col])

    df = df.sort_values('date')

    # Ensure numeric columns
    numeric_cols = ['Close', 'High', 'Low', 'Volume', 'Open', 'close', 'high', 'low', 'volume', 'open']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df.rename(columns={
        'close': 'Close', 'high': 'High', 'low': 'Low', 'volume': 'Volume',
        'open': 'Open', 'adj close': 'Adj Close'
    }, inplace=True)

    # Compute indicators
    df['Donchian_High'] = df['High'].rolling(window=20).max()
    df['Donchian_Low'] = df['Low'].rolling(window=20).min()
    df['Donchian_Mid'] = (df['Donchian_High'] + df['Donchian_Low']) / 2
    df['SMA20'] = df['Close'].rolling(window=20).mean()
    df['STD20'] = df['Close'].rolling(window=20).std()
    df['BB_Upper'] = df['SMA20'] + 2 * df['STD20']
    df['BB_Lower'] = df['SMA20'] - 2 * df['STD20']
    df['H-L'] = df['High'] - df['Low']
    df['H-PC'] = abs(df['High'] - df['Close'].shift(1))
    df['L-PC'] = abs(df['Low'] - df['Close'].shift(1))
    df['TR'] = df[['H-L', 'H-PC', 'L-PC']].max(axis=1)
    df['ATR'] = df['TR'].rolling(window=14).mean()
    df['Return'] = df['Close'].pct_change()
    df['Momentum'] = df['Close'] - df['Close'].shift(5)
    df['Volatility'] = df['Close'].rolling(window=5).std()
    df['OBV'] = (np.sign(df['Close'].diff()) * df['Volume']).fillna(0).cumsum()
    df['Volume_MA'] = df['Volume'].rolling(window=20).mean()

        # Additional Indicators (RSI and MACD)
    df['RSI'] = 100 - (100 / (1 + df['Close'].diff().rolling(14).mean() / df['Close'].diff().rolling(14).std()))
    df['EMA12'] = df['Close'].ewm(span=12, adjust=False).mean()
    df['EMA26'] = df['Close'].ewm(span=26, adjust=False).mean()
    df['MACD'] = df['EMA12'] - df['EMA26']
    df['Signal_Line'] = df['MACD'].ewm(span=9, adjust=False).mean()


    # Target and trends
    df['Predicted_Percentage_Change'] = (df['Close'].shift(-1) - df['Close']) / df['Close'] * 100
    df['Target'] = np.where(df['Predicted_Percentage_Change'] > 1, 1, 0)
    df['Volume_Trend'] = df['Volume'].rolling(window=5).mean() > df['Volume'].shift(5).rolling(window=5).mean()
    df['SMA_20D'] = df['Close'].rolling(window=20).mean()
    df['SMA_20W'] = df['Close'].rolling(window=100).mean()
    df['SMA_20M'] = df['Close'].rolling(window=100).mean()
    df['Daily_Up'] = df['Close'] > df['SMA_20D']
    df['Weekly_Down'] = df['Close'] < df['SMA_20W']
    df['Monthly_Down'] = df['Close'] < df['SMA_20M']
    df['Trend_Conflict'] = df['Daily_Up'] & (df['Weekly_Down'] | df['Monthly_Down'])

        # Validated Buy and Sell Signals
    df['Validated_Buy'] = (df['Target'] == 1) & (df['RSI'] > 30) & (df['Close'] > df['SMA20'])
    df['Validated_Sell'] = (df['Target'] == 0) & (df['RSI'] < 70) & (df['Close'] < df['SMA20'])

    features = ['Donchian_High', 'Donchian_Low', 'Donchian_Mid', 'SMA20', 'STD20', 'BB_Upper', 'BB_Lower',
            'ATR', 'Return', 'Momentum', 'Volatility', 'OBV', 'Volume_MA', 'RSI', 'MACD', 'Signal_Line']
    
    df.dropna(subset=features + ['Target'], inplace=True)

    if len(df) > 10 and df['Target'].nunique() > 1:

        X = df[features]
        y = df['Target']

        # Standardize features
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)



        # Check class distribution before applying SMOTE
        class_counts = Counter(y)
        minority_class_count = min(class_counts.values())
        if minority_class_count >= 6:
            smote = SMOTE(random_state=42)
            X_resampled, y_resampled = smote.fit_resample(X_scaled, y)
            #TotalStocksTrained += 1
        else:
            X_resampled, y_resampled = X_scaled, y

        # Train-test split
        X_train, X_test, y_train, y_test = train_test_split(X_resampled, y_resampled, test_size=0.2, random_state=42)

        # Train XGBoost model
        model = XGBClassifier(use_label_encoder=False, eval_metric='logloss')
        model.fit(X_train, y_train)

        # Predictions
        y_pred = model.predict(X_test)
        y_prob = model.predict_proba(X_test)[:, 1]

        df['Buy_Prob'] = model.predict_proba(X_scaled)[:, 1]
        df['Sell_Prob'] = model.predict_proba(X_scaled)[:, 0]
        df['ML_Prediction'] = model.predict(X_scaled)
        last_prediction = y_pred[-1]  # Last element in the prediction array
        # Evaluation
        accuracy = accuracy_score(y_test, y_pred)
        precision = precision_score(y_test, y_pred, zero_division=0)
        recall = recall_score(y_test, y_pred, zero_division=0)
        if len(np.unique(y_test)) > 1:
            roc_auc = roc_auc_score(y_test, y_prob)
        else:
            roc_auc = None
            #print("ROC AUC not defined: only one class present in y_test.")


        #print(f"Accuracy: {accuracy:.2%}")
        #print(f"Precision: {precision:.2%}")
        #print(f"Recall: {recall:.2%}")
        #if len(np.unique(y_test)) > 1:
            #print(f"ROC-AUC: {roc_auc:.2%}")

        # --- Predict Tomorrow ---
        last_row = X.iloc[[-1]].copy()


        tomorrow_features = last_row.copy()
        
        if not isinstance(df.index, pd.DatetimeIndex):
            df.index = pd.to_datetime(df['date'])

        tomorrow_features.index = [df.index[-1] + pd.Timedelta(days=1)]

        # Predict tomorrow's signal
        tomorrow_buy_prob = model.predict_proba(tomorrow_features)[:, 1][0]
        tomorrow_sell_prob = model.predict_proba(tomorrow_features)[:, 0][0]
        tomorrow_prediction = model.predict(tomorrow_features)[0]

        # Create tomorrow's prediction row
        df_tomorrow = pd.DataFrame({
            'date': [tomorrow_features.index[0]],
            'Buy_Prob': [tomorrow_buy_prob],
            'Sell_Prob': [tomorrow_sell_prob],
            'ML_Prediction': [tomorrow_prediction],
            'Strong_Buy': [(tomorrow_prediction == 1) and (tomorrow_buy_prob > strong_buy_threshold)],
            'Buy': [(tomorrow_prediction == 1) and (tomorrow_buy_prob > buy_threshold)]
        })

        high_conf_buy = df[(df['ML_Prediction'] == 1) & 
                           (df['Buy_Prob'] > strong_buy_threshold) & 
                           (df['Volume_Trend']) & 
                           (df['Validated_Buy']) & last_prediction]  # Added Validated_Buy as a check

        high_conf_sell = df[(df['ML_Prediction'] == 0) & (df['Sell_Prob'] > strong_sell_threshold)]

        conf_buy = df[(df['ML_Prediction'] == 1) & (df['Buy_Prob'] > buy_threshold)]
        conf_sell = df[(df['ML_Prediction'] == 0) & (df['Sell_Prob'] > sell_threshold)]

        filtered_buy = high_conf_buy[~high_conf_buy['Trend_Conflict']]
        filtered_sell = high_conf_sell[~high_conf_sell['Trend_Conflict']]

        return df, filtered_buy, filtered_sell, conf_buy, conf_sell, df_tomorrow, last_prediction
    else:
        #print("Not enough data after preprocessing to train the model.")
        return df, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# STATE MANAGEMENT WITH YOUR ORIGINAL ML LOGIC
# ==========================================

def run_discovery_mode():
    """Run at 9:30 AM - Discover new buy signals using your original logic"""
    print(">> MODE: DISCOVERY (9:30 AM) - Analyzing full market with ORIGINAL ML logic")
    
    results = []
    count = 0
    FailureCount = 0
    
    for ticker in nse_tickers:
        print(f"Analyzing {ticker}...")
        df = fetch_latest_data(ticker, period="1y", interval="1d")
        if df.empty:
            print(f"  ✗ Skipping {ticker} - no data")
            continue

        # Use your EXACT original analysis logic
        df, filtered_buy, _, conf_buy, conf_sell, df_tomorrow, today_prediction = analyze_stock(df)

        if not df.empty:

            try:
                latest_row = df.iloc[[-1]]  # Select the last row
                #print(latest_row)
                high_conf_buy = latest_row[
                (latest_row['ML_Prediction'] == 1) &
                (latest_row['Buy_Prob'] > buy_threshold) 
                #(latest_row['Volume_Trend']) &
                #(latest_row['Validated_Buy'])
                ]
            except KeyError as e:
                #print(f"Error: {e}")
                #print(latest_row.columns)
                FailureCount+=1
                #print(FailureCount)
                high_conf_buy = False
        else:
            count+=1
            #print("The DataFrame is empty. No rows to select.",count)
            high_conf_buy = False

        if isinstance(high_conf_buy, pd.DataFrame) and not high_conf_buy.empty:
            if 'ML_Prediction' in high_conf_buy.columns and (high_conf_buy['ML_Prediction'] == True).any():
                # Proceed with logic
                # Rest of the code
                results.append((
                    ticker,
                    df_tomorrow['Buy_Prob'].iloc[0],
                    df['Close'].iloc[-1],
                    df_tomorrow['ML_Prediction'].iloc[0]
                ))

    # Your original sorting logic - but now only including proper buy signals
    top_buys = sorted(results, key=lambda x: x[1], reverse=True)[:50]
    
    print(f"\n{'='*60}")
    print("DISCOVERY SUMMARY")
    print(f"{'='*60}")
    print(f"Stocks analyzed: {len(nse_tickers)}")
    print(f"Valid buy signals found: {len(results)}")
    print(f"Top buys selected: {len(top_buys)}")
    
    # Debug: Show probability distribution
    if results:
        probs = [r[1] for r in results]
        print(f"Buy probability range: {min(probs):.4f} - {max(probs):.4f}")
        print(f"Average buy probability: {np.mean(probs):.4f}")
    
    # Save to trading log - ONLY for valid buy signals
    if top_buys:
        df_log = load_trading_log()
        today_str = get_ist_time().strftime('%Y-%m-%d')
        
        new_entries = []
        for ticker, prob, price, pred, signal_type in top_buys:
            # Only add STRONG_BUY and BUY signals (not weak ones)
            if signal_type in ["STRONG_BUY", "BUY"]:
                new_entries.append({
                    'Date': today_str,
                    'Ticker': ticker,
                    'Buy_Price': price,
                    'Current_Price': price,
                    'Sell_Price': 0.0,
                    'PnL': 0.0,
                    'Status': 'OPEN',
                    'Buy_Prob': prob,
                    'Signal_Type': signal_type,
                    'Sell_Signal_Time': '',
                    'Investment_Amount': 100000,  # 1 lakh per stock
                    'Profit_Amount': 0.0,
                    'Return_Percentage': 0.0
                })
                print(f"  ➕ Adding to portfolio: {ticker} ({signal_type}, Prob: {prob:.4f})")
        
        if new_entries:
            new_df = pd.DataFrame(new_entries)
            df_log = pd.concat([df_log, new_df], ignore_index=True)
            save_trading_log(df_log)
        else:
            print("  ⚠️  No valid buy signals met probability thresholds")
        
        # Your original print output
        print("\nTop Buy Signals Selected for Tomorrow:")
        for ticker, prob, price, pred, signal_type in top_buys:
            if signal_type in ["STRONG_BUY", "BUY"]:
                print(f"  {ticker}: {signal_type}, Probability = {prob:.4f}, Today's Close = {price:.2f}")
    else:
        print("No valid buy signals found today (none met probability thresholds).")
    
    print(f"{'='*60}")

def calculate_profits(df_log):
    """Calculate today's profit and overall profit with accuracy metrics"""
    today_str = get_ist_time().strftime('%Y-%m-%d')
    
    # Today's trades
    today_trades = df_log[df_log['Date'] == today_str]
    today_closed = today_trades[today_trades['Status'] == 'CLOSED']
    
    # Overall trades
    all_closed = df_log[df_log['Status'] == 'CLOSED']
    all_trades = df_log
    
    # Profit calculations
    today_profit = today_closed['Profit_Amount'].sum() if not today_closed.empty else 0
    overall_profit = all_closed['Profit_Amount'].sum() if not all_closed.empty else 0
    
    # Accuracy calculations
    today_profitable = today_closed[today_closed['Profit_Amount'] > 0]
    overall_profitable = all_closed[all_closed['Profit_Amount'] > 0]
    
    today_accuracy = len(today_profitable) / len(today_closed) * 100 if len(today_closed) > 0 else 0
    overall_accuracy = len(overall_profitable) / len(all_closed) * 100 if len(all_closed) > 0 else 0
    
    return {
        'today_profit': today_profit,
        'overall_profit': overall_profit,
        'today_accuracy': today_accuracy,
        'overall_accuracy': overall_accuracy,
        'today_profitable_trades': len(today_profitable),
        'today_total_trades': len(today_closed),
        'overall_profitable_trades': len(overall_profitable),
        'overall_total_trades': len(all_closed),
        'today_trades_suggested': len(today_trades),
        'overall_trades_suggested': len(all_trades)
    }

def send_daily_summary_email(df_log, sender_email, recipients):
    """Send comprehensive daily summary email with profit tracking and accuracy"""
    today_str = get_ist_time().strftime('%Y-%m-%d')
    ist = ZoneInfo("Asia/Kolkata")
    now_ist = datetime.now(ist)
    
    # Calculate all profit metrics
    profit_metrics = calculate_profits(df_log)
    
    # Filter today's data
    today_data = df_log[df_log['Date'] == today_str]
    
    # Prepare data for email
    closed_trades = today_data[today_data['Status'] == 'CLOSED']
    open_trades = today_data[today_data['Status'] == 'OPEN']
    
    # Build HTML content
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 5px; }}
            .profit-positive {{ color: #27ae60; font-weight: bold; }}
            .profit-negative {{ color: #e74c3c; font-weight: bold; }}
            .profit-neutral {{ color: #7f8c8d; font-weight: bold; }}
            .accuracy-high {{ color: #27ae60; font-weight: bold; }}
            .accuracy-medium {{ color: #f39c12; font-weight: bold; }}
            .accuracy-low {{ color: #e74c3c; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
            th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
            th {{ background-color: #f8f9fa; }}
            .section {{ margin: 30px 0; }}
            .section-title {{ font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #2c3e50; }}
            .metric-card {{ background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 10px 0; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📊 Daily Trading Summary</h1>
            <p>Generated: {now_ist.strftime('%Y-%m-%d %H:%M IST')}</p>
        </div>
        
        <div class="section">
            <div class="section-title">📈 Prediction Accuracy</div>
            <div class="metric-card">
                <strong>Today's Accuracy:</strong><br>
                <span class="{'accuracy-high' if profit_metrics['today_accuracy'] > 70 else 'accuracy-medium' if profit_metrics['today_accuracy'] > 50 else 'accuracy-low'}">
                    {profit_metrics['today_profitable_trades']}/{profit_metrics['today_total_trades']} 
                    ({profit_metrics['today_accuracy']:.1f}%)
                </span><br>
                <small>{profit_metrics['today_trades_suggested']} stocks suggested today</small>
            </div>
            <div class="metric-card">
                <strong>Overall Accuracy:</strong><br>
                <span class="{'accuracy-high' if profit_metrics['overall_accuracy'] > 70 else 'accuracy-medium' if profit_metrics['overall_accuracy'] > 50 else 'accuracy-low'}">
                    {profit_metrics['overall_profitable_trades']}/{profit_metrics['overall_total_trades']} 
                    ({profit_metrics['overall_accuracy']:.1f}%)
                </span><br>
                <small>{profit_metrics['overall_trades_suggested']} total stocks suggested</small>
            </div>
        </div>

        <div class="section">
            <div class="section-title">💰 Profit Summary (₹1 Lakh per stock)</div>
            <table>
                <tr>
                    <td><strong>Today's Profit:</strong></td>
                    <td class="{'profit-positive' if profit_metrics['today_profit'] > 0 else 'profit-negative' if profit_metrics['today_profit'] < 0 else 'profit-neutral'}">
                        ₹{profit_metrics['today_profit']:,.2f}
                    </td>
                </tr>
                <tr>
                    <td><strong>Overall Profit:</strong></td>
                    <td class="{'profit-positive' if profit_metrics['overall_profit'] > 0 else 'profit-negative' if profit_metrics['overall_profit'] < 0 else 'profit-neutral'}">
                        ₹{profit_metrics['overall_profit']:,.2f}
                    </td>
                </tr>
                <tr>
                    <td><strong>Total Trades Today:</strong></td>
                    <td>{len(today_data)}</td>
                </tr>
                <tr>
                    <td><strong>Closed Trades Today:</strong></td>
                    <td>{len(closed_trades)}</td>
                </tr>
                <tr>
                    <td><strong>Open Trades (Carry Forward):</strong></td>
                    <td>{len(open_trades)}</td>
                </tr>
            </table>
        </div>
    """
    
    # Add Closed Trades section
    if not closed_trades.empty:
        html_body += """
        <div class="section">
            <div class="section-title">✅ Closed Trades Today</div>
            <table>
                <thead>
                    <tr>
                        <th>Ticker</th>
                        <th>Buy Price</th>
                        <th>Sell Price</th>
                        <th>P&L</th>
                        <th>Profit Amount</th>
                        <th>Return %</th>
                        <th>Sell Time</th>
                    </tr>
                </thead>
                <tbody>
        """
        for _, trade in closed_trades.iterrows():
            return_pct = (trade['PnL'] / trade['Buy_Price']) * 100
            pnl_class = "profit-positive" if trade['PnL'] > 0 else "profit-negative" if trade['PnL'] < 0 else "profit-neutral"
            html_body += f"""
                    <tr>
                        <td>{trade['Ticker']}</td>
                        <td>₹{trade['Buy_Price']:.2f}</td>
                        <td>₹{trade['Sell_Price']:.2f}</td>
                        <td class="{pnl_class}">₹{trade['PnL']:.2f}</td>
                        <td class="{pnl_class}">₹{trade['Profit_Amount']:,.2f}</td>
                        <td class="{pnl_class}">{return_pct:.2f}%</td>
                        <td>{trade['Sell_Signal_Time']}</td>
                    </tr>
            """
        html_body += """
                </tbody>
            </table>
        </div>
        """
    
    # Add Open Trades section
    if not open_trades.empty:
        html_body += """
        <div class="section">
            <div class="section-title">⏳ Open Trades (Carrying Forward)</div>
            <table>
                <thead>
                    <tr>
                        <th>Ticker</th>
                        <th>Buy Price</th>
                        <th>Current Price</th>
                        <th>Current P&L</th>
                        <th>Unrealized Profit</th>
                        <th>Return %</th>
                    </tr>
                </thead>
                <tbody>
        """
        for _, trade in open_trades.iterrows():
            current_pnl = trade['Current_Price'] - trade['Buy_Price']
            unrealized_profit = trade['Profit_Amount']  # Already calculated in save_trading_log
            return_pct = (current_pnl / trade['Buy_Price']) * 100
            pnl_class = "profit-positive" if current_pnl > 0 else "profit-negative" if current_pnl < 0 else "profit-neutral"
            html_body += f"""
                    <tr>
                        <td>{trade['Ticker']}</td>
                        <td>₹{trade['Buy_Price']:.2f}</td>
                        <td>₹{trade['Current_Price']:.2f}</td>
                        <td class="{pnl_class}">₹{current_pnl:.2f}</td>
                        <td class="{pnl_class}">₹{unrealized_profit:,.2f}</td>
                        <td class="{pnl_class}">{return_pct:.2f}%</td>
                    </tr>
            """
        html_body += """
                </tbody>
            </table>
        </div>
        """
    
    html_body += """
    </body>
    </html>
    """
    
    # Build the email
    msg = EmailMessage()
    msg["Subject"] = f"Daily Trading Summary - {today_str} | Accuracy: {profit_metrics['today_accuracy']:.1f}% | Profit: ₹{profit_metrics['today_profit']:,.2f}"
    msg["From"] = formataddr(("Quant Signals Bot", sender_email))
    msg["To"] = ", ".join(recipients)
    msg.add_alternative(html_body, subtype="html")
    
    # Send email
    username = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASS")
    
    if not username or not password:
        print("SMTP credentials missing")
        return False
        
    try:
        # Auto-detect SMTP server
        email_domain = sender_email.split('@')[1].lower() if '@' in sender_email else ''
        smtp_server = 'smtp.gmail.com' if 'gmail' in email_domain else 'smtp.office365.com'
        
        with smtplib.SMTP(smtp_server, 587, timeout=30) as server:
            server.set_debuglevel(0)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(username, password)
            server.send_message(msg)
        print("✓ Daily summary email sent successfully!")
        return True
    except Exception as e:
        print(f"✗ Failed to send daily summary email: {e}")
        return False
        
def run_monitoring_mode():
    """Run at 10:00 AM, 10:30 AM, etc. - Monitor existing positions"""
    print(">> MODE: MONITORING - Tracking existing positions with ORIGINAL ML logic")
    
    df_log = load_trading_log()
    today_str = get_ist_time().strftime('%Y-%m-%d')
    current_time_str = get_ist_time().strftime('%H:%M')
    
    # Filter for today's OPEN trades
    open_trades = df_log[(df_log['Date'] == today_str) & (df_log['Status'] == 'OPEN')]
    
    if open_trades.empty:
        print("No open positions to monitor today.")
        return
    
    print(f"Monitoring {len(open_trades)} open positions...")
    
    updates = []
    
    for idx, trade in open_trades.iterrows():
        ticker = trade['Ticker']
        buy_price = trade['Buy_Price']
        
        print(f"Checking {ticker}...")
        
        # Get live price
        live_price = fetch_live_price(ticker)
        if live_price is None:
            print(f"  ✗ Could not fetch live price for {ticker}")
            continue
        
        # Run your original analysis to check for sell signals
        df = fetch_latest_data(ticker, period="1y", interval="1d")
        if df.empty:
            print(f"  ✗ No data for {ticker}")
            continue
        
        try:
            # Use your EXACT original analysis logic
            df, filtered_buy, _, conf_buy, conf_sell, df_tomorrow, today_prediction = analyze_stock(df)
            
            if df.empty:
                print(f"  ✗ Analysis failed for {ticker}")
                continue
            
            # Update current price and PnL
            df_log.at[idx, 'Current_Price'] = live_price
            current_pnl = live_price - buy_price
            df_log.at[idx, 'PnL'] = current_pnl
            
            # Use your original sell logic
            sell_triggered = False
            reason = ""
            
            # Check if the model now predicts sell (class 0)
            tomorrow_prediction = df_tomorrow['ML_Prediction'].iloc[0] if not df_tomorrow.empty else None
            tomorrow_sell_prob = df_tomorrow['Sell_Prob'].iloc[0] if not df_tomorrow.empty else 0
            
            if tomorrow_prediction == 0:
                sell_triggered = True
                reason = "Model Flip to Sell"
            elif tomorrow_sell_prob > strong_sell_threshold:
                sell_triggered = True
                reason = f"High Sell Prob ({tomorrow_sell_prob:.4f})"
            # Optional: Add stop loss
            elif current_pnl / buy_price < -0.02:  # 2% stop loss
                sell_triggered = True
                reason = "Stop Loss Triggered"
            
            if sell_triggered:
                print(f"  ✓ SELL Signal: {ticker} - {reason}")
                df_log.at[idx, 'Status'] = 'CLOSED'
                df_log.at[idx, 'Sell_Price'] = live_price
                df_log.at[idx, 'Sell_Signal_Time'] = current_time_str
                
                updates.append({
                    'Ticker': ticker,
                    'Action': 'SELL',
                    'Price': live_price,
                    'PnL': f"{current_pnl:.2f} ({current_pnl/buy_price*100:.2f}%)",
                    'Reason': reason
                })
            else:
                print(f"  ✓ HOLD: {ticker}, Live: {live_price:.2f}, PnL: {current_pnl:.2f}")
                updates.append({
                    'Ticker': ticker,
                    'Action': 'HOLD', 
                    'Price': live_price,
                    'PnL': f"{current_pnl:.2f}",
                    'Reason': 'No sell signal'
                })
                
        except Exception as e:
            print(f"  ✗ Error monitoring {ticker}: {e}")
    
    # Save updates
    if save_trading_log(df_log):
        print("✓ Trading log updated")
    

def main():
    """Main execution with state management and daily summary"""
    today_str = get_ist_time().strftime('%Y-%m-%d')
    current_time_str = get_ist_time().strftime('%H:%M')
    current_hour = get_ist_time().hour
    current_minute = get_ist_time().minute
    
    print(f"\n{'='*80}")
    print(f"QUANT SIGNALS BOT - {today_str} {current_time_str} IST")
    print(f"{'='*80}")
    
    # Load trading log to determine mode
    df_log = load_trading_log()
    todays_trades = df_log[df_log['Date'] == today_str]
    
    # Check if this is the last run of the day (3:00 PM)
    is_last_run = (current_hour == 15 and current_minute == 0)
    
    # AUTO-RESET: If market is closed (after 3:30 PM IST), reset for next day
    # if current_hour > 15 or (current_hour == 15 and current_minute >= 30):
    #     print(">> Market closed - resetting for next day...")
    #     df_log = df_log[df_log['Date'] != today_str]
    #     save_trading_log(df_log)
    #     todays_trades = pd.DataFrame()
    #     print("✓ Reset complete - ready for discovery mode tomorrow")
    #     return
    
    if todays_trades.empty:
        # DISCOVERY MODE: No trades for today yet
        print(">> MODE: DISCOVERY")
        run_discovery_mode()
    else:
        # MONITORING MODE: Trades already exist for today
        print(">> MODE: MONITORING")
        run_monitoring_mode()
    
    # Send daily summary only at the last run (3:00 PM)
    if is_last_run:
        print(">> SENDING DAILY SUMMARY REPORT")
        df_log = load_trading_log()  # Reload to get latest data
        recipients_env = os.getenv("MAIL_RECIPIENTS", "")
        recipients = [r.strip() for r in recipients_env.split(",") if r.strip()]
        sender_email = os.getenv("SMTP_USER")
        
        if sender_email and recipients:
            send_daily_summary_email(df_log, sender_email, recipients)
        else:
            print("✗ Cannot send daily summary: Missing email configuration")
    
    print(f"\n{'='*80}")
    print("EXECUTION COMPLETE")
    print(f"{'='*80}")
    
if __name__ == "__main__":
    main()
