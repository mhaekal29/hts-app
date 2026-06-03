// === HOETRANGSA v4 — BAGIAN 1 DARI 3 ===
import { useState, useEffect, useCallback, useRef, useMemo, createContext, useContext } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area, ReferenceLine } from "recharts";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
var SIZES=["5.5 kg","12 kg","50 kg"];
var JENIS=["Isi","Tabung+Isi"];
var DENOMS=[100000,50000,20000,10000,5000,2000,1000];
// Urutan role untuk dropdown
var ROLE_ORDER=["admin","akuntan","sales_driver","sales_freelance","checker","driver_truck","owner"];
function sortEmp(emps){return(emps||[]).slice().sort((a,b)=>{var ia=ROLE_ORDER.indexOf(a.role);var ib=ROLE_ORDER.indexOf(b.role);if(ia!==ib)return ia-ib;return(a.nama||"").localeCompare(b.nama||"");});}
// Load Plus Jakarta Sans font
(function(){if(typeof document!=="undefined"&&!document.getElementById("pjs-font")){var l=document.createElement("link");l.id="pjs-font";l.rel="stylesheet";l.href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap";document.head.appendChild(l);}})();
var UANG_MAKAN_DEFAULT=15000;
var SPBE_LOC={"SPPBE KCR":25000,"SPPBE MGL":115000};
var KAT_K=["BBM","Gaji","Uang Makan Karyawan","Perbaikan","Administrasi","Sewa","Parkir","Service Kendaraan","Fee LPG 50kg","Uang Jalan SPBE","Uang Bongkar DO","Pancung 12kg","Pancung 5.5kg","Listrik PLN","Internet Wi-Fi","Air Galon","Lainnya"];
var KAT_AUTO_HARGA={"Uang Bongkar DO":50000,"Pancung 12kg":200000,"Pancung 5.5kg":95000};
var SPPBE_OPTS=["SPPBE KCR","SPPBE MGL","Lainnya"];
var PLG_KAT_27=["Rumah Tangga","Restaurant","Cafe","Warung Kopi","Warung Nasi/Rumah Makan","Kedai Bakso/Mie Ayam","Kantin Sekolah/Kampus","Catering","SPPG","UMKM Kuliner","King Fried Chicken","Toko Roti","Pabrik Roti","Produksi Industri Makanan","Peternakan Ayam","Hotel","Rumah Sakit","Kantor Pemerintah","Kantor Swasta","Kampus/Universitas","Pesantren","Pangkalan","Agen LPG","Reseller LPG","Sub Agen LPG","Canvaser LPG","SPBU","Swalayan/Grosir","Mini Market/Kios","WO/Event Organizer","Usaha Umum","Laundry","Industri Rumahan","Lainnya"]
var PLG_TITIP_KAT=PLG_KAT_27.filter(k=>!["Rumah Tangga","Industri Rumahan"].includes(k));
var KAT_OPS=["BBM","Uang Bongkar DO","Uang Jalan SPBE","Fee LPG 50kg","Pancung 12kg","Pancung 5.5kg","Parkir","Service Kendaraan","Lainnya","Listrik PLN","Internet Wi-Fi","Air Galon"];
var ABSENSI_STATUS=["Hadir","Sakit","Izin","Alpha","Cuti","Libur"];
var ROLE_LBL={admin:"Manajer",akuntan:"Kasir/Akuntan",sales_driver:"Sales Driver",sales_freelance:"Sales Freelance",checker:"Checker",driver_truck:"Driver Truck",owner:"Komisaris"};
var ROLE_TABS={owner:null,admin:null,akuntan:null,sales_driver:["dashboard","penjualan","piutang","setoran","pelanggan"],sales_freelance:["dashboard","penjualan","piutang","setoran","pelanggan"],checker:["dashboard","stok","do"],driver_truck:["dashboard","do"]};
var PENJUALAN_ROLES=["admin","akuntan","sales_driver","sales_freelance","owner"];
var DEF_MODAL={"Isi":{"5.5 kg":98881,"12 kg":212133,"50 kg":1026464},"Tabung+Isi":{"5.5 kg":381645,"12 kg":730990,"50 kg":2415419}};
var DEF_HET={"Isi":{"5.5 kg":110000,"12 kg":230000,"50 kg":1200000},"Tabung+Isi":{"5.5 kg":395000,"12 kg":730000,"50 kg":2450000}};
var PIE_COLORS=["#0C4DA2","#00A651","#F39C12","#ED1C24","#8E44AD","#1ABC9C","#E67E22","#3498DB"];
var ROMAN=["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];
var BULAN_ID=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

var DEF_EMP=[
{id:"e01",username:"Haekal",password:"Sudirman80",role:"admin",nama:"Muhammad Haekal",posisi:"Manajer",telepon:"082246980311",alamat:"Banda Aceh",gajiPokok:5000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e02",username:"DAYAT",password:"Sudirman80",role:"akuntan",nama:"MANARUL HIDAYAT",posisi:"Kasir/Akuntan",telepon:"082277214045",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e03",username:"MUSLEM",password:"Sudirman80",role:"sales_driver",nama:"MUSLEM",posisi:"Sales Driver",telepon:"082259494802",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e04",username:"USNI",password:"Sudirman80",role:"sales_driver",nama:"USNI SIREGAR",posisi:"Sales Driver",telepon:"082187102114",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e05",username:"SAYUTI",password:"Sudirman80",role:"sales_driver",nama:"T. SAYUTI",posisi:"Sales Driver",telepon:"081377494144",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e06",username:"HELMI",password:"Sudirman80",role:"sales_driver",nama:"HELMI ZAKARIA",posisi:"Sales Driver",telepon:"081377190552",alamat:"Lam Lumpu - Peukan Bada, Aceh Besar",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e07",username:"HAMDANI",password:"Sudirman80",role:"driver_truck",nama:"HAMDANI IBRAHIM",posisi:"Driver Truck SPBE",telepon:"085276590894",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e08",username:"BUNYAMIN",password:"Sudirman80",role:"sales_freelance",nama:"BUNYAMIN",posisi:"Sales Marketing",telepon:"085212570213",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e09",username:"SAIBAN",password:"Sudirman80",role:"checker",nama:"IBNU SAIBAN",posisi:"Checker",telepon:"082298704089",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e10",username:"SYAHRULI",password:"Sudirman80",role:"owner",nama:"SYAHRULI",posisi:"Komisaris",telepon:"081370492714",alamat:"",gajiPokok:2500000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e11",username:"FRIDZLUN",password:"Sudirman80",role:"sales_freelance",nama:"FRIDZLUN",posisi:"Sales Freelance",telepon:"085361224830",alamat:"",gajiPokok:0,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e12",username:"Bachtiar",password:"Sudirman80",role:"sales_freelance",nama:"Bachtiar",posisi:"Sales Freelance",telepon:"081264600500",alamat:"",gajiPokok:0,uangMakan:15000,uangMakanMode:"harian",aktif:true},
];

// ─── GOOGLE SHEETS SYNC ───────────────────────────────────────────────────────
var GAS_URL="https://script.google.com/macros/s/AKfycbypH6BaonwpselzUFZ3Q0QJHvRbTinzpgSR37aJpKZIjX_8XVvjPC2tK1CFi8Gst7RRwg/exec";
var GAS_SECRET="HTS2026";
var SYNC_TABLES=["penjualan","bon","pengeluaran","pelanggan","employees","stok","doList","absensi","payrollLog","ambilan","titipList","setoranLog","tutupBuku","config"];

// CORS-safe: kirim via form POST ke GAS
function gasPost(payload){
  return new Promise(function(resolve,reject){
    var id="cb_"+Date.now();
    // Gunakan no-cors fetch — tidak bisa baca response, tapi data terkirim
    // Untuk read: pakai JSONP via script tag
    fetch(GAS_URL,{
      method:"POST",
      mode:"no-cors",
      headers:{"Content-Type":"text/plain"},
      body:JSON.stringify(payload)
    }).then(()=>resolve({ok:true})).catch(e=>reject(e));
  });
}

// JSONP untuk READ (bisa baca response)
function gasReadJsonp(table){
  return new Promise(function(resolve){
    var cbName="hts_cb_"+Date.now()+"_"+Math.random().toString(36).slice(2);
    window[cbName]=function(data){
      delete window[cbName];
      if(s)document.head.removeChild(s);
      resolve(data&&data.ok?data.data:[]);
    };
    var s=document.createElement("script");
    var params=new URLSearchParams({secret:GAS_SECRET,action:"read",table,callback:cbName});
    s.src=GAS_URL+"?"+params.toString();
    s.onerror=function(){delete window[cbName];resolve([]);};
    document.head.appendChild(s);
    setTimeout(function(){if(window[cbName]){delete window[cbName];resolve([]);}},15000);
  });
}

async function gasRead(table){
  try{return await gasReadJsonp(table);}catch(e){return[];}
}

async function gasWrite(table,data){
  try{
    await gasPost({secret:GAS_SECRET,action:"write",table,data});
    return true;
  }catch(e){return false;}
}

// Push semua data ke Google Sheets
async function pushAll(data,setSyncStatus){
  setSyncStatus("pushing");
  try{
    var tasks=[];
    tasks.push(gasWrite("penjualan",data.penjualan||[]));
    tasks.push(gasWrite("bon",data.bon||[]));
    tasks.push(gasWrite("pengeluaran",data.pengeluaran||[]));
    tasks.push(gasWrite("pelanggan",data.pelanggan||[]));
    tasks.push(gasWrite("employees",data.employees||[]));
    tasks.push(gasWrite("doList",data.doList||[]));
    tasks.push(gasWrite("absensi",data.absensi||[]));
    tasks.push(gasWrite("payrollLog",data.payrollLog||[]));
    tasks.push(gasWrite("ambilan",data.ambilan||[]));
    tasks.push(gasWrite("titipList",data.titipList||[]));
    tasks.push(gasWrite("setoranLog",data.setoranLog||[]));
    tasks.push(gasWrite("tutupBuku",data.tutupBuku||[]));
    // stok & config sebagai object → wrap dalam array
    tasks.push(gasWrite("stok",[{key:"stok",val:JSON.stringify({stock:data.stock,totalAsset:data.totalAsset,stockLog:data.stockLog,modalHistory:data.modalHistory,hetPrices:data.hetPrices,counters:data.counters,theme:data.theme})}]));
    tasks.push(gasWrite("config",[{key:"company",val:JSON.stringify(data.company||{})}]));
    await Promise.all(tasks);
    setSyncStatus("ok");
    return true;
  }catch(e){setSyncStatus("error");return false;}
}

// Pull semua data dari Google Sheets
async function pullAll(setSyncStatus){
  setSyncStatus("pulling");
  try{
    var [penj,bon,pen,plg,emp,doL,abs,pay,amb,titip,setor,tb,stokRaw,confRaw]=await Promise.all([
      gasRead("penjualan"),gasRead("bon"),gasRead("pengeluaran"),gasRead("pelanggan"),
      gasRead("employees"),gasRead("doList"),gasRead("absensi"),gasRead("payrollLog"),
      gasRead("ambilan"),gasRead("titipList"),gasRead("setoranLog"),gasRead("tutupBuku"),
      gasRead("stok"),gasRead("config")
    ]);
    var stokMeta={};
    if(stokRaw&&stokRaw.length>0){try{stokMeta=JSON.parse(stokRaw[0].val);}catch(e){}}
    var company={};
    if(confRaw&&confRaw.length>0){try{company=JSON.parse(confRaw[0].val);}catch(e){}}
    setSyncStatus("ok");
    return{penjualan:penj,bon,pengeluaran:pen,pelanggan:plg,
      employees:emp.length>0?emp:null,
      doList:doL,absensi:abs,payrollLog:pay,ambilan:amb,
      titipList:titip,setoranLog:setor,tutupBuku:tb,
      company,
      ...stokMeta};
  }catch(e){setSyncStatus("error");return null;}
}

var INIT={
stock:{"5.5 kg":0,"12 kg":0,"50 kg":0},stokKosong:{"5.5 kg":0,"12 kg":0,"50 kg":0},totalAsset:{"5.5 kg":0,"12 kg":0,"50 kg":0},
stockLog:[],doList:[],modalHistory:[],hetPrices:{},titipList:[],
penjualan:[],bon:[],pengeluaran:[],employees:DEF_EMP.slice(),
tutupBuku:[],pelanggan:[],setoranSales:[],setoranLog:[],absensi:[],ambilan:[],payrollLog:[],
counters:{inv:{},sg:{},reg:0},theme:"light",
company:{nama:"PT. HOE TRANG SA",alamat:"Jl. Jendral Sudirman No.80, Geuce Iniem, Kec. Banda Raya, Kota Banda Aceh",telepon:"0812 6900 2121",telepon2:"(0651) 21221",email:"npso.pthoetrangsa@gmail.com",website:"pt-hoetrangsa.business.site",npwp:"",slogan:"DEALER LPG PERTAMINA",bankNama:"BSI",bankAtasNama:"PT. HOE TRANG SA",bankRekening:"812 69 2121 8",logo:"",logoPertamina:"",ttdKasir:"",ttdDirektur:"",stempelLunas:"",direkturNama:"Muhammad Haekal",kasirNama:"MANARUL HIDAYAT",soldTo:"731070",shipToKCR:"862070",shipToMGL:"782092",saBulan:"Juni 2026",sa12KCR:"2845075",sa55KCR:"2845111",sa12MGL:"",sa55MGL:"",assetArmada:0,hargaTbgKosong:{"5.5 kg":0,"12 kg":0,"50 kg":0}}
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,5);}
function toDay(){return new Date().toISOString().split("T")[0];}
function toMonth(){return toDay().slice(0,7);}
function fD(d){return!d?"-":new Date(d+"T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});}
function fDs(d){return!d?"-":new Date(d+"T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"});}
function fR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
function dLeft(ds){if(!ds)return null;var t=new Date();t.setHours(0,0,0,0);return Math.ceil((new Date(ds+"T00:00:00")-t)/86400000);}
function iTotal(its){return its.reduce((a,it)=>a+Number(it.qty||0)*Number(it.price||0),0);}
function daysInMonth(ym){var[y,m]=ym.split("-").map(Number);return new Date(y,m,0).getDate();}
function safeFileName(s){return(s||"file").replace(/[\/\\?%*:|"<>]/g,"_").replace(/\s+/g,"_").slice(0,80);}
function getModal(data,ukuran,jenis,tgl){tgl=tgl||toDay();var h=(data.modalHistory||[]).filter(x=>x.ukuran===ukuran&&x.jenis===jenis&&x.tanggal<=tgl);if(!h.length)return(DEF_MODAL[jenis]||{})[ukuran]||0;return h.slice().sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0].harga;}
function getHET(data,ukuran,jenis){var hp=data.hetPrices;if(hp&&hp[jenis]&&hp[jenis][ukuran]!=null)return hp[jenis][ukuran];return(DEF_HET[jenis]||{})[ukuran]||0;}
function calcMargin(items,data,tgl){return items.reduce((a,it)=>{var m=getModal(data,it.ukuran,it.jenis,tgl);return a+(Number(it.price||0)-m)*Number(it.qty||0);},0);}
function getKonsumenTitipBal(titipList){var bal={};(titipList||[]).forEach(t=>{var k=t.konsumenNama;if(!k)return;if(!bal[k])bal[k]={"5.5 kg":0,"12 kg":0,"50 kg":0,telp:t.konsumenTelp||"",alamat:t.konsumenAlamat||""};var items=t.items&&t.items.length>0?t.items:(t.ukuran&&t.qty?[{ukuran:t.ukuran,qty:t.qty}]:[]);var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;items.forEach(it=>{if(it.ukuran&&bal[k][it.ukuran]!==undefined)bal[k][it.ukuran]=m>0?(bal[k][it.ukuran]||0)+Number(it.qty||0):Math.max(0,(bal[k][it.ukuran]||0)-Number(it.qty||0));});});return bal;}
function getTitipTotal(titipList,ukuran){var bal=getKonsumenTitipBal(titipList);return Object.values(bal).reduce((a,v)=>a+Math.max(0,v[ukuran]||0),0);}
function getKosong(data,ukuran){return Math.max(0,((data.stokKosong||{})[ukuran]||0));}

function terbilang(n){
if(n==null||isNaN(n))return"Nol";n=Math.floor(Math.abs(n));if(n===0)return"Nol";
var s=["","Satu","Dua","Tiga","Empat","Lima","Enam","Tujuh","Delapan","Sembilan","Sepuluh","Sebelas"];
function f(n){
if(n<12)return s[n];
if(n<20)return f(n-10)+" Belas";
if(n<100)return f(Math.floor(n/10))+" Puluh"+(n%10?" "+f(n%10):"");
if(n<200)return"Seratus"+(n-100?" "+f(n-100):"");
if(n<1000)return f(Math.floor(n/100))+" Ratus"+(n%100?" "+f(n%100):"");
if(n<2000)return"Seribu"+(n-1000?" "+f(n-1000):"");
if(n<1000000)return f(Math.floor(n/1000))+" Ribu"+(n%1000?" "+f(n%1000):"");
if(n<1000000000)return f(Math.floor(n/1000000))+" Juta"+(n%1000000?" "+f(n%1000000):"");
return f(Math.floor(n/1000000000))+" Milyar"+(n%1000000000?" "+f(n%1000000000):"");}
return f(n).trim()+" Rupiah";}

function nextInvNo(data,tanggal){var d=new Date(tanggal+"T00:00:00");var m=d.getMonth();var y=d.getFullYear();var key=y+"-"+String(m+1).padStart(2,"0");var ctr=(data.counters?.inv||{})[key]||0;var n=ctr+1;return{no:"#HTS/INV/"+ROMAN[m]+"."+String(y).slice(-2)+"/"+String(n).padStart(3,"0"),key,n};}
function nextSGNo(data,bulan){var[y,m]=bulan.split("-").map(Number);var ctr=(data.counters?.sg||{})[bulan]||0;var n=ctr+1;return{no:String(n).padStart(3,"0")+"/HTS/SG/"+ROMAN[m-1]+"/"+y,key:bulan,n};}
function nextRegNo(data){var ctr=(data.counters?.reg||0)+1;return{no:"HTS/CST/"+String(ctr).padStart(3,"0"),n:ctr};}

function calcBonusSales(empId,bulan,data){var penj=(data.penjualan||[]).filter(p=>p.salesId===empId&&(p.tanggal||"").startsWith(bulan));var q55=0,q12=0;penj.forEach(p=>(p.items||[]).forEach(it=>{if(it.ukuran==="5.5 kg")q55+=Number(it.qty||0);else if(it.ukuran==="12 kg")q12+=Number(it.qty||0);}));var r55=q55<500?500:q55;var b55=q55*r55;var r12=q12<500?500:q12;var b12=q12*r12;return{q55,r55,b55,q12,r12,b12,total:b55+b12};}

function getBiayaOpsAuto(empId,bulan,data){return(data.pengeluaran||[]).filter(p=>p.karyawanId===empId&&(p.tanggal||"").startsWith(bulan)&&KAT_OPS.includes(p.kategori)).map(p=>({id:p.id,label:p.kategori+(p.ket?" - "+p.ket:""),nominal:Number(p.nominal||0),kategori:p.kategori}));}

function getPinjamanSaldo(empId,bulan,data){var amb=(data.ambilan||[]).filter(a=>a.karyawanId===empId&&(a.tanggal||"")<=(bulan+"-31"));var totAmb=amb.reduce((a,x)=>a+Number(x.nominal||0),0);var pots=(data.payrollLog||[]).filter(p=>p.empId===empId&&p.bulan<bulan);var totPot=pots.reduce((a,x)=>a+Number(x.potonganPinjaman||0),0);return Math.max(0,totAmb-totPot);}

function calcPayrollFull(emp,bulan,data){var abs=(data.absensi||[]).filter(a=>a.karyawanId===emp.id&&(a.tanggal||"").startsWith(bulan));var hariHadir=abs.filter(a=>a.status==="Hadir").length;var totalHariKerja=daysInMonth(bulan);var absen=abs.filter(a=>["Alpha"].includes(a.status)).length;var isSales=["sales_driver","sales_freelance"].includes(emp.role);var bonus=isSales?calcBonusSales(emp.id,bulan,data):{q55:0,r55:500,b55:0,q12:0,r12:500,b12:0,total:0};var mode=emp.uangMakanMode||"harian";var makanAuto=(data.pengeluaran||[]).filter(p=>p.karyawanId===emp.id&&["Uang Makan Karyawan","Uang Makan","Makan Karyawan"].includes(p.kategori)&&(p.tanggal||"").startsWith(bulan)).reduce((a,p)=>a+Number(p.nominal||0),0);var uangMakan=mode==="harian"?makanAuto:hariHadir*(emp.uangMakan||UANG_MAKAN_DEFAULT);var ops=getBiayaOpsAuto(emp.id,bulan,data).filter(x=>!["Uang Makan","Makan Karyawan"].some(k=>x.label.startsWith(k)));var bongkarTotal=ops.filter(x=>x.kategori==="Uang Bongkar DO").reduce((a,x)=>a+x.nominal,0);var spbbeTotal=ops.filter(x=>x.kategori==="Uang Jalan SPBE").reduce((a,x)=>a+x.nominal,0);var bongkarCount=ops.filter(x=>x.kategori==="Uang Bongkar DO").length;var spbbeCount=ops.filter(x=>x.kategori==="Uang Jalan SPBE").length;var pinjamanSaldo=getPinjamanSaldo(emp.id,bulan,data);return{hariHadir,totalHariKerja,absen,gajiPokok:emp.gajiPokok||0,bonus,uangMakanMode:mode,uangMakan,bongkarTotal,spbbeTotal,bongkarCount,spbbeCount,pinjamanSaldo};}

function calcNilaiStok(data){return SIZES.reduce((a,s)=>{var qty=(data.stock||{})[s]||0;var modal=getModal(data,s,"Isi");return a+qty*modal;},0);}
function calcPinjamanKaryawan(data){return(data.ambilan||[]).reduce((a,x)=>a+Number(x.nominal||0),0)-(data.payrollLog||[]).reduce((a,x)=>a+Number(x.potonganPinjaman||0),0);}
function calcTotalPiutang(data){return(data.bon||[]).filter(b=>b.status!=="lunas").reduce((a,b)=>a+(b.sisaTagihan||0),0);}

// ─── PRINT HELPER (Simpan PDF via dialog browser) ─────────────────────────────
// Tidak butuh CDN. User pilih "Save as PDF" di dialog print browser.

function doPrint(id){
var e=document.getElementById("__pst");if(e)e.remove();
var st=document.createElement("style");st.id="__pst";
st.textContent=[
"@media print{",
"  *{visibility:hidden!important;margin:0!important;}",
"  #"+id+",#"+id+" *{visibility:visible!important;}",
"  #"+id+"{",
"    position:fixed!important;inset:0!important;",
"    padding:0!important;margin:0!important;",
"    background:white!important;",
"    z-index:99999!important;",
"    overflow:visible!important;",
"    box-shadow:none!important;",
"    border-radius:0!important;",
"  }",
"  @page{margin:10mm;size:A4;}",
"}"
].join("");
document.head.appendChild(st);
window.print();
setTimeout(()=>{var e=document.getElementById("__pst");if(e)e.remove();},2000);
}

// ─── THEME ────────────────────────────────────────────────────────────────────
var THEMES={
dark:{bg:"#0B1520",card:"#111E2D",bdr:"#1A2D42",nav:"#1A2B3C",red:"#C0392B",rdk:"#7B241C",rlt:"#E74C3C",org:"#E67E22",olt:"#F39C12",grn:"#1D8348",glt:"#27AE60",blu:"#1A5276",blt:"#2980B9",gry:"#4A5568",gl2:"#718096",gltr:"#A0AEC0",wht:"#F7FAFC",mode:"dark",inHv:"#0A2040",inHvE:"#4A0E0E"},
light:{bg:"#F1F5F9",card:"#FFFFFF",bdr:"#E2E8F0",nav:"#F8FAFC",red:"#DC2626",rdk:"#991B1B",rlt:"#EF4444",org:"#EA580C",olt:"#F97316",grn:"#15803D",glt:"#16A34A",blu:"#0C4DA2",blt:"#2563EB",gry:"#94A3B8",gl2:"#64748B",gltr:"#475569",wht:"#0F172A",mode:"light",inHv:"#DBEAFE",inHvE:"#FEE2E2"}
};
var ThemeCtx=createContext(THEMES.light);
function useTheme(){return useContext(ThemeCtx);}

// ─── HOOKS ────────────────────────────────────────────────────────────────────
function useMobile(){var[m,setM]=useState(()=>typeof window!=="undefined"&&window.innerWidth<680);useEffect(()=>{var fn=()=>setM(window.innerWidth<680);window.addEventListener("resize",fn);return()=>window.removeEventListener("resize",fn);},[]);return m;}
function useToast(){var[toasts,setToasts]=useState([]);function toast(msg,type="success"){var id=uid();setToasts(t=>[...t,{id,msg,type}]);setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3000);}return{toasts,toast};}

// ─── UI PRIMITIVES ────────────────────────────────────────────────────────────
function Toast({toasts}){var C=useTheme();if(!toasts.length)return null;return <div style={{position:"fixed",top:70,right:16,zIndex:9999,display:"flex",flexDirection:"column",gap:8}}>{toasts.map(t=><div key={t.id} style={{background:t.type==="error"?C.rdk:t.type==="warning"?C.org:C.grn,color:"#FFF",padding:"10px 16px",borderRadius:8,fontSize:13,fontWeight:600,boxShadow:"0 4px 20px rgba(0,0,0,.3)",maxWidth:320}}>{t.msg}</div>)}</div>;}
function Bdg({color="gray",children}){var C=useTheme();var d=C.mode==="dark";var m=d?{red:["#4A0E0E",C.rlt],green:["#0A2E14",C.glt],orange:["#3D200A",C.olt],blue:["#0A2040",C.blt],gray:["#1E2A35",C.gl2]}:{red:["#FEE2E2","#991B1B"],green:["#DCFCE7","#166534"],orange:["#FED7AA","#9A3412"],blue:["#DBEAFE","#1E40AF"],gray:["#E2E8F0","#475569"]};var p=m[color]||m.gray;return <span style={{background:p[0],color:p[1],padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{children}</span>;}
function Card({children,style={}}){var C=useTheme();return <div style={{background:C.card,borderRadius:12,border:"1px solid "+C.bdr,padding:16,marginBottom:12,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.04)":"none",...style}}>{children}</div>;}
function STitle({icon,children}){var C=useTheme();return <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}><span style={{fontSize:22}}>{icon}</span><h2 style={{margin:0,fontSize:17,fontWeight:800,color:C.wht}}>{children}</h2></div>;}
function Inp({label,type="text",value,onChange,placeholder="",ro=false,style={}}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10,...style}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input type={type} value={value!=null?value:""} readOnly={ro} placeholder={placeholder} onChange={e=>onChange&&onChange(e.target.value)} style={{...IS,background:ro?C.bg:C.nav}}/></div>;}
function Sel({label,value,onChange,opts=[],style={}}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10,...style}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<select value={value!=null?value:""} onChange={e=>onChange(e.target.value)} style={{...IS,background:C.nav}}>{opts.map(o=>{var v=o!=null&&o.v!=null?o.v:o;var l=o!=null&&o.l!=null?o.l:o;return <option key={String(v)} value={String(v)}>{String(l)}</option>;})}</select></div>;}
function Btn({children,color="blue",sm=false,dis=false,style={},onClick}){var C=useTheme();var bgs={blue:[C.blu,C.blt],red:[C.rdk,C.red],green:[C.grn,C.glt],orange:[C.mode==="dark"?"#854000":"#C2410C",C.org],gray:[C.gry,C.gl2]};var[hov,setHov]=useState(false);var pair=bgs[color]||bgs.blue;return <button onClick={onClick} disabled={dis} onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)} style={{background:dis?C.gry:hov?pair[1]:pair[0],color:"#FFF",border:"none",borderRadius:8,padding:sm?"6px 10px":"9px 16px",fontSize:sm?11:13,fontWeight:700,cursor:dis?"not-allowed":"pointer",transition:"all .15s",...style}}>{children}</button>;}
function Tbl({headers=[],rows=[],empty="Belum ada data"}){var C=useTheme();return <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><tr>{headers.map((h,i)=><th key={i} style={{padding:"8px 10px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"left",fontSize:11,borderBottom:"2px solid "+C.bdr,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead><tbody>{!rows.length?<tr><td colSpan={headers.length} style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</td></tr>:rows.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>{r.map((c,j)=><td key={j} style={{padding:"7px 10px",color:C.wht,verticalAlign:"middle"}}>{c}</td>)}</tr>)}</tbody></table></div>;}
function MCards({headers,rows,empty="Belum ada data"}){var C=useTheme();if(!rows.length)return <div style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</div>;return <div>{rows.map((r,i)=><div key={i} style={{background:C.nav,borderRadius:10,padding:12,marginBottom:8,border:"1px solid "+C.bdr}}>{headers.map((h,j)=><div key={j} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",borderBottom:j<headers.length-1?"1px solid "+C.bdr:"none"}}><span style={{fontSize:11,color:C.gl2,marginRight:8}}>{h}</span><span style={{fontSize:12,color:C.wht,textAlign:"right"}}>{r[j]}</span></div>)}</div>)}</div>;}
function RTbl({headers,rows,empty}){var m=useMobile();return m?<MCards headers={headers} rows={rows} empty={empty}/>:<Tbl headers={headers} rows={rows} empty={empty}/>;}
function SC({label,value,icon,color,sub}){var C=useTheme();return <div style={{background:C.card,borderRadius:12,padding:"12px 14px",border:"1px solid "+C.bdr,flex:1,minWidth:110,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.04)":"none"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div><div style={{fontSize:10,color:C.gl2,fontWeight:600,marginBottom:2}}>{label}</div><div style={{fontSize:15,fontWeight:900,color:color||C.blt,lineHeight:1.1}}>{value}</div>{sub&&<div style={{fontSize:10,color:C.gl2,marginTop:2}}>{sub}</div>}</div><span style={{fontSize:18,opacity:.7}}>{icon}</span></div></div>;}
function AutoInp({label,value="",onChange,options=[],placeholder="",onSelect}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};var[show,setShow]=useState(false);var filtered=options.filter(o=>!value||o.toLowerCase().includes(value.toLowerCase())).slice(0,7);return <div style={{position:"relative",marginBottom:10}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input value={value} placeholder={placeholder} onChange={e=>{onChange(e.target.value);setShow(true);}} onFocus={()=>setShow(true)} onBlur={()=>setTimeout(()=>setShow(false),200)} style={{...IS,background:C.nav}}/>{show&&filtered.length>0&&<div style={{position:"absolute",top:"100%",left:0,right:0,background:C.card,border:"1px solid "+C.bdr,borderRadius:8,zIndex:200,maxHeight:160,overflowY:"auto",boxShadow:"0 4px 16px rgba(0,0,0,.3)"}}>{filtered.map(o=><div key={o} onMouseDown={()=>{onChange(o);if(onSelect)onSelect(o);setShow(false);}} style={{padding:"9px 12px",cursor:"pointer",color:C.wht,fontSize:13,borderBottom:"1px solid "+C.bdr}}>{o}</div>)}</div>}</div>;}
function ActBtns({onEdit,onDel,onPrint}){var C=useTheme();return <div style={{display:"flex",gap:4}}>{onPrint&&<button onClick={onPrint} title="Cetak Invoice" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button>}{onEdit&&<button onClick={onEdit} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>✏️</button>}{onDel&&<button onClick={onDel} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button>}</div>;}
function ConfirmDel({msg,onConfirm,onCancel}){var C=useTheme();return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:8500,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{background:C.card,borderRadius:14,padding:28,maxWidth:340,width:"90%",border:"1px solid "+C.rlt}}><div style={{color:C.wht,fontWeight:700,marginBottom:8,fontSize:16}}>🗑️ Hapus Data?</div><div style={{color:C.gl2,fontSize:13,marginBottom:20}}>{msg}</div><div style={{display:"flex",gap:10}}><Btn color="red" onClick={onConfirm}>Ya, Hapus</Btn><Btn color="gray" onClick={onCancel}>Batal</Btn></div></div></div>;}
function Modal({title,children,onClose,onSave,saveLabel="💾 Simpan",width=560}){var C=useTheme();return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:8000,display:"flex",alignItems:"center",justifyContent:"center",padding:12}}><div style={{background:C.card,borderRadius:14,padding:22,maxWidth:width,width:"100%",border:"1px solid "+C.bdr,maxHeight:"92vh",overflowY:"auto"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><b style={{color:C.wht,fontSize:15}}>{title}</b><button onClick={onClose} style={{background:"none",border:"none",color:C.gl2,fontSize:22,cursor:"pointer"}}>✕</button></div>{children}{onSave&&<div style={{display:"flex",gap:10,marginTop:14}}><Btn color="green" onClick={onSave}>{saveLabel}</Btn><Btn color="gray" onClick={onClose}>Batal</Btn></div>}</div></div>;}
function MonthPicker({label,value,onChange}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input type="month" value={value} onChange={e=>onChange(e.target.value)} style={{...IS,background:C.nav}}/></div>;}

// ─── FILTER TABLE (NEW v4) ────────────────────────────────────────────────────
// Tabel dengan: filter mini di header + sort klik header, dipakai di Penjualan, Piutang, Pengeluaran, Laporan
function FilterTbl({columns,data:rows,empty="Belum ada data",maxRows=200,renderMobileCard}){
var C=useTheme();var mob=useMobile();
var[hdrFilters,setHdrFilters]=useState({});
var[sort,setSort]=useState({key:null,dir:"asc"});
var filtered=useMemo(()=>{
return rows.filter(row=>{
return columns.every(col=>{
if(!col.filterable)return true;
var fv=(hdrFilters[col.key]||"").trim().toLowerCase();
if(!fv)return true;
var val=col.filterVal?col.filterVal(row):row[col.key];
if(val==null)return false;
return String(val).toLowerCase().includes(fv);
});
});
},[rows,columns,hdrFilters]);
var sorted=useMemo(()=>{
if(!sort.key)return filtered;
var col=columns.find(c=>c.key===sort.key);if(!col)return filtered;
var get=col.sortVal||(r=>r[sort.key]);
var arr=filtered.slice();
arr.sort((a,b)=>{var va=get(a);var vb=get(b);if(va==null)return 1;if(vb==null)return -1;if(typeof va==="number"&&typeof vb==="number")return sort.dir==="asc"?va-vb:vb-va;return sort.dir==="asc"?String(va).localeCompare(String(vb)):String(vb).localeCompare(String(va));});
return arr;
},[filtered,sort,columns]);
var display=sorted.slice(0,maxRows);
function toggleSort(k){var col=columns.find(c=>c.key===k);if(!col||col.sortable===false)return;setSort(s=>s.key===k?{key:k,dir:s.dir==="asc"?"desc":"asc"}:{key:k,dir:"asc"});}
function setFilter(k,v){setHdrFilters(p=>({...p,[k]:v}));}
if(mob&&renderMobileCard){return <div>{display.map((r,i)=><div key={i}>{renderMobileCard(r,i)}</div>)}{!display.length&&<div style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</div>}</div>;}
return <div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
<thead>
<tr>{columns.map(c=><th key={c.key} onClick={()=>toggleSort(c.key)} style={{padding:"8px 10px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"left",fontSize:11,borderBottom:"2px solid "+C.bdr,whiteSpace:"nowrap",cursor:c.sortable!==false?"pointer":"default",userSelect:"none",width:c.width?(typeof c.width==="number"?c.width+"px":c.width):undefined,minWidth:c.width?(typeof c.width==="number"?c.width+"px":c.width):undefined}}>{c.label}{c.sortable!==false&&<span style={{marginLeft:4,color:sort.key===c.key?C.blt:C.gry,fontSize:9}}>{sort.key===c.key?(sort.dir==="asc"?"▲":"▼"):"⇅"}</span>}</th>)}</tr>
<tr>{columns.map(c=><th key={c.key+"_f"} style={{padding:"4px 6px",background:C.bg,borderBottom:"1px solid "+C.bdr}}>{c.filterable===false?<span/>:c.filterType==="select"?<select value={hdrFilters[c.key]||""} onChange={e=>setFilter(c.key,e.target.value)} style={{width:"100%",background:C.card,border:"1px solid "+C.bdr,borderRadius:5,padding:"3px 5px",fontSize:11,color:C.wht,outline:"none"}}><option value="">Semua</option>{(c.options||[]).map(o=><option key={o.v||o} value={o.v||o}>{o.l||o}</option>)}</select>:<input value={hdrFilters[c.key]||""} onChange={e=>setFilter(c.key,e.target.value)} placeholder="Filter..." style={{width:"100%",background:C.card,border:"1px solid "+C.bdr,borderRadius:5,padding:"3px 6px",fontSize:11,color:C.wht,outline:"none"}}/>}</th>)}</tr>
</thead>
<tbody>
{!display.length?<tr><td colSpan={columns.length} style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</td></tr>:display.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>{columns.map(c=><td key={c.key} style={{padding:"7px 10px",color:C.wht}}>{c.render?c.render(r):r[c.key]}</td>)}</tr>)}
</tbody>
</table>
</div>
{(Object.values(hdrFilters).some(v=>v)||sort.key)&&<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",fontSize:11,color:C.gl2}}><span>{sorted.length} dari {rows.length} hasil</span><Btn sm color="gray" onClick={()=>{setHdrFilters({});setSort({key:null,dir:"asc"});}}>✕ Reset Filter</Btn></div>}
</div>;
}

// ─── LOGO ─────────────────────────────────────────────────────────────────────
function LPGLogo({size=72}){return <svg width={size} height={size*0.85} viewBox="0 0 70 60" fill="none"><rect x="2" y="14" width="13" height="40" fill="#ED1C24" rx="1"/><rect x="19" y="2" width="13" height="52" fill="#F2A900" rx="1"/><rect x="36" y="10" width="13" height="44" fill="#00A651" rx="1"/><text x="26" y="14" textAnchor="middle" fontSize="7" fontWeight="700" fill="#0A2C5C">1982</text></svg>;}
function CompanyLogo({company={},h=52}){if(company.logo)return <img src={company.logo} style={{height:h,objectFit:"contain"}} alt="logo"/>;return <LPGLogo size={h}/>;}
function PertaminaLogo({company={},h=32}){if(company.logoPertamina)return <img src={company.logoPertamina} style={{height:h,objectFit:"contain"}} alt="pertamina"/>;return <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{display:"flex",height:h*0.85,transform:"skewX(-12deg)"}}><div style={{width:h*0.5,background:"#ED1C24"}}/><div style={{width:h*0.4,background:"#00A651"}}/><div style={{width:h*0.4,background:"#0C4DA2"}}/></div><span style={{fontSize:h*0.65,fontWeight:800,color:"#00713A",letterSpacing:0.5}}>PERTAMINA</span></div>;}

// === AKHIR BAGIAN 1 ===
// === BAGIAN 2 DARI 3 ===

// ─── INVOICE BARU v4 (no double logo, PDF/PNG buttons) ────────────────────────
function InvoiceView({inv,company={},onClose}){
if(!inv)return null;
var isLunas=inv.metodeBayar!=="BON"&&(inv.metodeBayar||"").toLowerCase().indexOf("bon")<0&&!inv.isBon;
var total=inv.total||0;
var NAVY="#0a1f44";var NAVY2="#122d5e";var BLUE="#1565c0";
var RED="#e53935";var GREEN="#6ab04c";var TEAL="#00acc1";var TEAL_LIGHT="#e0f7fa";
var WHITE="#fff";var G100="#eef2f9";var G200="#dde4f0";var G400="#8fa3c0";var G600="#4a6080";
var FONT="'Plus Jakarta Sans',-apple-system,'Segoe UI',sans-serif";
var HARI_ID=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
function fDHari(d){if(!d)return"-";var dt=new Date(d+"T00:00:00");return HARI_ID[dt.getDay()]+", "+dt.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});}
var salesNama=inv.salesNama||"";
return <div id="_inv_wrap" style={{position:"fixed",inset:0,background:"#cdd3db",zIndex:9500,padding:16,overflowY:"auto"}}>
<div id="_inv" style={{maxWidth:700,margin:"0 auto",background:WHITE,borderRadius:10,overflow:"hidden",boxShadow:"0 20px 60px rgba(10,31,68,.14)",color:"#0d1f3c",fontFamily:FONT}}>

{/* HEADER navy */}
<div style={{background:"linear-gradient(135deg,"+NAVY+" 0%,"+NAVY2+" 100%)",padding:"20px 26px 17px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{display:"flex",alignItems:"center",gap:12}}>
{company.logo
?<img src={company.logo} style={{height:54,objectFit:"contain"}} alt="logo"/>
:<div style={{display:"flex",alignItems:"center",gap:12}}>
<div style={{display:"flex",gap:3,alignItems:"flex-end"}}>
<div style={{width:9,height:28,background:RED,borderRadius:2}}/>
<div style={{width:9,height:20,background:"#1e88e5",borderRadius:2}}/>
<div style={{width:9,height:28,background:GREEN,borderRadius:2}}/>
</div>
<div>
<div style={{fontSize:18,fontWeight:800,color:WHITE,letterSpacing:.3,lineHeight:1.1}}>{company.nama||"PT. HOE TRANG SA"}</div>
<div style={{fontSize:9,fontWeight:600,color:RED,letterSpacing:1.8,textTransform:"uppercase",marginTop:3}}>{company.slogan||"DEALER LPG PERTAMINA"}</div>
</div>
</div>}
</div>
{/* Pertamina */}
<div style={{display:"flex",alignItems:"center",gap:7}}>
{company.logoPertamina
?<img src={company.logoPertamina} style={{height:38,objectFit:"contain"}} alt="pertamina"/>
:<div style={{display:"flex",alignItems:"center",gap:7}}><svg width={54} height={40} viewBox="0 0 60 42" fill="none">
<polygon points="18,0 32,14 18,28 4,14" fill="#E53935"/>
<polygon points="14,28 28,14 14,0 0,14" fill="#1565C0" opacity="0.85"/>
<polygon points="22,28 36,14 22,0 8,14" fill="#6AB04C" opacity="0.9"/>
</svg>
<span style={{fontSize:14,fontWeight:800,color:WHITE,letterSpacing:.4}}>PERTAMINA</span></div>}
</div>
</div>

{/* DIVIDER 3 warna */}
<div style={{height:3,display:"flex"}}>
<div style={{flex:1,background:BLUE}}/><div style={{flex:1,background:GREEN}}/><div style={{flex:1,background:RED}}/>
</div>

{/* TITLE BLOCK — putih */}
<div style={{background:WHITE,padding:"18px 26px 16px"}}>
{/* Judul + badge status */}
<div style={{display:"flex",alignItems:"center",justifyContent:"center",position:"relative",marginBottom:4}}>
<div style={{fontSize:26,fontWeight:800,color:NAVY,letterSpacing:3,textAlign:"center"}}>I N V O I C E</div>
{isLunas
?<div style={{position:"absolute",right:0,fontSize:10,fontWeight:700,padding:"3px 11px",borderRadius:20,background:"#D1FAE5",border:"1px solid #10B981",color:"#065F46"}}>✓ Lunas</div>
:<div style={{position:"absolute",right:0,fontSize:10,fontWeight:700,padding:"3px 11px",borderRadius:20,background:"#FEE2E2",border:"1px solid #EF4444",color:"#991B1B"}}>— Belum Lunas</div>}
</div>
{/* No invoice di bawah judul */}
<div style={{textAlign:"center",fontSize:11,fontWeight:600,color:G400,letterSpacing:.5,marginBottom:14}}>NO. <span style={{color:NAVY,fontWeight:700}}>{inv.noInv}</span></div>
{/* Kepada kiri, Tanggal+Sales kanan */}
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:16}}>
<div style={{borderLeft:"3px solid "+NAVY,paddingLeft:12}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1.2,textTransform:"uppercase",color:G400,marginBottom:4}}>Kepada Yth.</div>
<div style={{fontSize:17,fontWeight:800,color:NAVY,lineHeight:1.2,marginBottom:2}}>{inv.konsumen}</div>
<div style={{fontSize:11,color:G400}}>Di — {inv.kota||"Banda Aceh"}</div>
</div>
<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:5}}>
<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G400,marginBottom:1}}>Tanggal</div>
<div style={{fontSize:12,fontWeight:700,color:NAVY}}>{fDHari(inv.tanggal)}</div>
</div>
{salesNama&&<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G400,marginBottom:1}}>Sales</div>
<div style={{fontSize:11.5,fontWeight:700,color:BLUE}}>{salesNama}</div>
</div>}
</div>
</div>
</div>

{/* GAP abu */}
<div style={{height:12,background:"#f4f7fc"}}/>

{/* TABEL — putih */}
<div style={{margin:"0 26px",borderRadius:8,overflow:"hidden",border:"1px solid "+G200,boxShadow:"0 2px 12px rgba(10,31,68,.06)"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontFamily:FONT,fontSize:11.5}}>
<thead><tr style={{background:NAVY}}>
{[{l:"Date",w:"18%",a:"left"},{l:"Detail Produk",w:"26%",a:"left"},{l:"Qty",w:"10%",a:"center"},{l:"",w:"5%",a:"center"},{l:"Unit Price",w:"19%",a:"right"},{l:"Line Total",w:"22%",a:"right"}].map((h,i)=><th key={i} style={{width:h.w,padding:"9px 12px",color:WHITE,fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",textAlign:h.a}}>{h.l}</th>)}
</tr></thead>
<tbody>
{(inv.items||[]).map((it,i)=>{
var prodLabel=it.jenis==="Tabung+Isi"?"Tbg + Isi":"Refill";
return <tr key={i} style={{background:i%2===0?WHITE:G100}}>
<td style={{padding:"10px 14px",color:G400,fontSize:10,fontWeight:600,lineHeight:1.5,borderBottom:"1px solid "+G200,whiteSpace:"nowrap"}}>{fDHari(inv.tanggal).split(",")[0]}<br/>{fDs(inv.tanggal)}</td>
<td style={{padding:"10px 12px",borderBottom:"1px solid "+G200}}>
<span style={{display:"inline-block",background:NAVY,color:WHITE,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700,marginRight:6}}>{it.ukuran}</span>
<span style={{fontSize:10.5,color:G600,marginRight:3}}>{prodLabel}</span>
<span style={{fontWeight:600,color:"#0d1f3c"}}>LPG</span>
</td>
<td style={{padding:"10px 12px",textAlign:"center",fontWeight:700,color:NAVY,borderBottom:"1px solid "+G200}}>{it.qty}</td>
<td style={{padding:"10px 12px",textAlign:"center",color:G400,fontSize:11,fontWeight:700,borderBottom:"1px solid "+G200}}>×</td>
<td style={{padding:"10px 12px",textAlign:"right",borderBottom:"1px solid "+G200}}>
<span style={{fontSize:10,fontWeight:500,color:G400,marginRight:1}}>Rp </span>
<span style={{fontWeight:700,color:NAVY}}>{Number(it.price||0).toLocaleString("id-ID")}</span>
</td>
<td style={{padding:"10px 12px",textAlign:"right",borderBottom:"1px solid "+G200}}>
<span style={{fontSize:10,fontWeight:500,color:G400,marginRight:1}}>Rp </span>
<span style={{fontWeight:700,color:NAVY,fontSize:12.5}}>{(Number(it.qty||0)*Number(it.price||0)).toLocaleString("id-ID")}</span>
</td>
</tr>;})}
{Array.from({length:Math.max(0,2-(inv.items||[]).length)}).map((_,i)=><tr key={"e"+i} style={{background:((inv.items||[]).length+i)%2===0?WHITE:G100}}><td style={{padding:"8px 12px",borderBottom:"1px solid "+G200}}>&nbsp;</td><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/></tr>)}
</tbody>
</table>
{/* TERBILANG | LUNAS | GRAND TOTAL */}
<div style={{display:"flex",alignItems:"stretch",borderTop:"2px solid "+G200}}>
<div style={{flex:1,padding:"13px 14px",background:TEAL_LIGHT}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G600,marginBottom:4}}>Terbilang</div>
<div style={{fontSize:11.5,fontWeight:500,fontStyle:"italic",color:NAVY,lineHeight:1.5}}># {terbilang(total)} #</div>
</div>
{/* Stempel LUNAS di antara terbilang & grand total */}
<div style={{display:"flex",alignItems:"center",justifyContent:"center",padding:"0 10px",background:TEAL_LIGHT}}>
{isLunas&&(company.stempelLunas
?<img src={company.stempelLunas} style={{height:52,opacity:.9,transform:"rotate(-5deg)"}} alt="lunas"/>
:<div style={{fontSize:13,fontWeight:800,color:RED,border:"2px solid "+RED,padding:"5px 11px",borderRadius:3,letterSpacing:4,opacity:.88,transform:"rotate(-5deg)",whiteSpace:"nowrap"}}>L U N A S</div>)}
</div>
<div style={{minWidth:190,padding:"13px 16px",background:WHITE,borderLeft:"2px solid "+G200,display:"flex",flexDirection:"column",justifyContent:"center",alignItems:"flex-end"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G400,marginBottom:4}}>Grand Total</div>
<div style={{display:"flex",alignItems:"baseline",gap:4}}>
<span style={{fontSize:13,fontWeight:600,color:G400}}>Rp</span>
<span style={{fontSize:24,fontWeight:800,color:NAVY,letterSpacing:.3}}>{Number(total).toLocaleString("id-ID")}</span>
</div>
</div>
</div>
</div>

{/* GAP abu */}
<div style={{height:4,background:"#f4f7fc"}}/>

{/* BOTTOM — kiri rekening+WA, kanan TTD */}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,padding:"16px 26px 18px",background:WHITE}}>
<div style={{display:"flex",flexDirection:"column",gap:10}}>

{/* REKENING — teal-light, bold, no rek besar */}
<div style={{borderRadius:8,overflow:"hidden",border:"1px solid "+G200}}>
<div style={{background:NAVY2,padding:"8px 14px"}}>
<div style={{fontSize:13,fontWeight:800,color:WHITE,letterSpacing:.5}}>BANK {company.bankNama||"BSI"}</div>
</div>
<div style={{background:TEAL_LIGHT}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid rgba(0,172,193,.15)"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:2}}>A/C Atas Nama</div>
<div style={{fontSize:14,fontWeight:800,color:NAVY}}>{company.bankAtasNama||company.nama}</div>
</div>
<div style={{padding:"8px 14px"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:2}}>No. Rekening</div>
<div style={{fontFamily:"'Courier New',Courier,monospace",fontSize:20,fontWeight:800,color:NAVY,letterSpacing:1.5}}>{company.bankRekening||"812 69 2121 8"}</div>
</div>
</div>
</div>

{/* WA Konfirmasi */}
<div style={{background:TEAL_LIGHT,borderRadius:8,border:"1px solid "+G200,padding:"10px 13px"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:6}}>Konfirmasi Bukti Pembayaran Ke</div>
<div style={{fontSize:22,fontWeight:800,color:NAVY,lineHeight:1,letterSpacing:.5}}>{company.telepon||"0812 6900 2121"}</div>
</div>

<div style={{fontSize:10,color:G600,fontStyle:"italic",lineHeight:1.5}}>Demikian kami sampaikan, atas perhatian dan kerjasamanya kami ucapkan terima kasih.</div>
</div>

{/* TTD kanan — kotak persegi ramping */}
<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-start",paddingTop:8,gap:4,textAlign:"center"}}>
<div style={{fontSize:10,color:G400}}>Banda Aceh, {fDHari(inv.tanggal)}</div>
<div style={{fontSize:11,fontWeight:700,color:NAVY,marginTop:1,marginBottom:10}}>{company.nama}</div>
<div style={{width:140,height:70,border:"1.5px dashed "+G200,borderRadius:8,background:WHITE,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:G400,fontStyle:"italic",marginBottom:10,lineHeight:1.4,textAlign:"center",flexShrink:0}}>
{company.ttdKasir?<img src={company.ttdKasir} style={{height:66,objectFit:"contain"}} alt="ttd"/>:<><div style={{fontSize:8,color:"#8fa3c0"}}>TTD Elektronik</div><div style={{fontSize:8,color:"#8fa3c0"}}>upload di Pengaturan</div></>}
</div>
<div style={{fontSize:14,fontWeight:800,color:NAVY,borderTop:"2px solid "+NAVY,paddingTop:5,display:"inline-block",minWidth:140}}>{company.kasirNama||"MANARUL HIDAYAT"}</div>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1.2,textTransform:"uppercase",color:G400,marginTop:3}}>Kasir</div>
</div>
</div>

{/* FOOTER navy */}
<div style={{background:NAVY,padding:"13px 22px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
{[["📍","Alamat",[company.alamat]],["📞","No.HP & WA",[company.telepon,"TELP. "+company.telepon2]],["🌐","Email & Website",[company.email,company.website]]].map((x,i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:10}}>
<div style={{width:28,height:28,background:"rgba(255,255,255,.1)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:13}}>{x[0]}</div>
<div><div style={{fontSize:7,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:"rgba(255,255,255,.4)",marginBottom:3}}>{x[1]}</div>{x[2].map((v,j)=><div key={j} style={{fontSize:9.5,fontWeight:600,color:WHITE,lineHeight:1.5}}>{v}</div>)}</div>
</div>)}
</div>

</div>
<div style={{maxWidth:700,margin:"12px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<button onClick={()=>doPrint("_inv")} style={{background:NAVY,color:WHITE,border:"none",padding:"10px 28px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>🖨️ Cetak / Simpan PDF</button>
<span style={{fontSize:11,color:"#888",alignSelf:"center",fontStyle:"italic"}}>Pilih "Save as PDF" di dialog print</span>
<button onClick={onClose} style={{background:"#566573",color:WHITE,border:"none",padding:"10px 22px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>✕ Tutup</button>
</div>
</div>;
}

// ─── KWITANSI SLIP GAJI BARU v4 (no O/X column, PDF/PNG) ──────────────────────
function SlipGajiView({slip,company={},onClose}){
if(!slip)return null;
var rows=slip.rows||[];
function rR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
var totPgh=rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totDiterima=totPgh-totPot-totYdt;
var sisa=Number(slip.totalPinjaman||0)-Number(slip.potonganPinjaman||0);
var blnIdx=slip.bulan?Number(slip.bulan.split("-")[1])-1:0;
var thIdx=slip.bulan?slip.bulan.split("-")[0]:"";
var fileName=safeFileName((slip.nama||"karyawan")+"_"+(slip.noSlip||"SLIP").replace(/[\/]/g,"-"));
var hasUploadedLogo=!!company.logo;
var penghasilanRows=rows.filter(r=>r.section==="penghasilan");
var potonganRows=rows.filter(r=>r.section==="potongan");
var ydtRows=rows.filter(r=>r.section==="ydt");
var SNAVY="#0a1f44";var SNAVY2="#122d5e";var SRED="#e53935";var SGREEN="#6ab04c";var SWHITE="#fff";var SFONT="'Plus Jakarta Sans',-apple-system,'Segoe UI',sans-serif";
return <div id="_slip_wrap" style={{position:"fixed",inset:0,background:"#cdd3db",zIndex:9700,padding:16,overflowY:"auto",fontFamily:SFONT}}>
<div id="_slip" style={{maxWidth:680,margin:"0 auto",background:SWHITE,borderRadius:10,overflow:"hidden",boxShadow:"0 20px 60px rgba(10,31,68,.14)",color:"#0d1f3c",fontSize:13}}>
{/* HEADER — sama dengan Invoice */}
<div style={{background:"linear-gradient(135deg,"+SNAVY+" 0%,"+SNAVY2+" 100%)",padding:"18px 24px 15px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{display:"flex",alignItems:"center",gap:12}}>
{company.logo
?<img src={company.logo} style={{height:48,objectFit:"contain"}} alt="logo"/>
:<div style={{display:"flex",alignItems:"center",gap:12}}>
<div style={{display:"flex",gap:3,alignItems:"flex-end"}}>
<div style={{width:8,height:26,background:SRED,borderRadius:2}}/>
<div style={{width:8,height:18,background:"#1e88e5",borderRadius:2}}/>
<div style={{width:8,height:26,background:SGREEN,borderRadius:2}}/>
</div>
<div>
<div style={{fontSize:17,fontWeight:800,color:SWHITE,letterSpacing:.3,lineHeight:1.1}}>{company.nama||"PT. HOE TRANG SA"}</div>
<div style={{fontSize:9,fontWeight:600,color:SRED,letterSpacing:1.8,textTransform:"uppercase",marginTop:3}}>{company.slogan||"DEALER LPG PERTAMINA"}</div>
</div>
</div>}
</div>
{/* Pertamina diamond */}
<div style={{display:"flex",alignItems:"center",gap:6}}>
{company.logoPertamina
?<img src={company.logoPertamina} style={{height:34,objectFit:"contain"}} alt="pertamina"/>
:<div style={{display:"flex",alignItems:"center",gap:6}}><svg width={48} height={36} viewBox="0 0 60 42" fill="none">
<polygon points="18,0 32,14 18,28 4,14" fill="#E53935"/>
<polygon points="14,28 28,14 14,0 0,14" fill="#1565C0" opacity="0.85"/>
<polygon points="22,28 36,14 22,0 8,14" fill="#6AB04C" opacity="0.9"/>
</svg>
<span style={{fontSize:13,fontWeight:800,color:SWHITE,letterSpacing:.4}}>PERTAMINA</span></div>}
</div>
</div>
{/* DIVIDER 3 warna */}
<div style={{height:3,display:"flex"}}>
<div style={{flex:1,background:"#1565c0"}}/><div style={{flex:1,background:"#6ab04c"}}/><div style={{flex:1,background:"#e53935"}}/>
</div>
<div style={{display:"flex",justifyContent:"space-between",padding:"8px 18px",borderBottom:"1px solid #ccc",fontSize:12.5}}>
<div><span style={{color:"#444"}}>No.Pembayaran :</span> <b>{slip.noSlip}</b></div>
<div><span style={{color:"#444"}}>Tanggal :</span> <b>{fDs(slip.tanggal)}</b></div>
</div>
<h2 style={{textAlign:"center",fontSize:20,fontWeight:800,padding:"10px 0 6px",letterSpacing:1,margin:0}}>KWITANSI SLIP GAJI {(BULAN_ID[blnIdx]||"").toUpperCase()} {thIdx}</h2>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",padding:"6px 18px 10px",gap:"4px 24px",fontSize:13.5}}>
<div><b style={{minWidth:64,display:"inline-block"}}>NAMA</b> : {slip.nama}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Alamat</b> : {slip.alamat||"-"}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Jabatan</b> : {slip.posisi}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Telepon</b> : {slip.telepon||"-"}</div>
</div>
<div style={{padding:"8px 18px",display:"flex",justifyContent:"space-between",borderTop:"1px solid #aaa",borderBottom:"1px solid #aaa",fontSize:13,background:"#FAF5E8"}}>
<span>Total Hari Kerja: <b>{slip.totalHariKerja}</b></span>
<span>Hadir: <b>{slip.hariHadir}</b></span>
<span style={{color:"#d63030"}}>Absen: <b>{slip.absen}</b></span>
</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
<thead><tr>
<th style={{width:36,padding:"8px 4px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"center"}}>NO</th>
<th style={{padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"left"}}>KETERANGAN</th>
<th style={{width:200,padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"center"}}>@</th>
<th style={{width:140,padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"right"}}>JUMLAH</th>
</tr></thead>
<tbody>
<tr><td colSpan={4} style={{background:"#EDF3FB",padding:"7px 10px"}}><b style={{textDecoration:"underline"}}>Penghasilan</b></td></tr>
{penghasilanRows.map((r,i)=><tr key={"p"+i} style={{background:"#EDF3FB"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#666"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,color:"#444",textAlign:"center"}}>{r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:600}}>{Number(r.jumlah||0)>0?rR(r.jumlah):"Rp -"}</td>
</tr>)}
<tr style={{background:"#EDF3FB",borderTop:"1px solid #333"}}><td colSpan={3} style={{padding:"7px 10px",textAlign:"right",fontWeight:800}}>Total Penghasilan</td><td style={{padding:"7px 10px",textAlign:"right",fontWeight:800}}>{rR(totPgh)} <span style={{color:"#00A651",fontWeight:900}}>+</span></td></tr>
<tr><td colSpan={4} style={{height:6,background:"white"}}/></tr>
{potonganRows.map((r,i)=><tr key={"pt"+i} style={{background:"#FDF6E1",color:"#a32d2d"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#a32d2d"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,fontStyle:"italic",textAlign:"center"}}>{r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:700}}>{r.kind==="info"?"":(Number(r.jumlah||0)>0?rR(r.jumlah)+" -":"Rp -")}</td>
</tr>)}
{Number(slip.totalPinjaman||0)>0&&<tr style={{background:"#FDF6E1"}}>
<td/>
<td style={{padding:"5px 10px",fontStyle:"italic",textDecoration:"underline",color:"#a32d2d",fontWeight:700}}>Sisa Pinjaman</td>
<td style={{padding:"5px 10px",textAlign:"center"}}><span style={{background:"#FFF35B",color:"#1a1a1a",padding:"2px 10px",borderRadius:3,border:"1px solid #d4af00",fontWeight:800}}>{rR(sisa)}</span></td>
<td/>
</tr>}
<tr style={{background:"#FDF6E1"}}>
<td colSpan={2} style={{textAlign:"right",fontWeight:800,color:"#a32d2d",padding:"6px 10px"}}>TOTAL POTONGAN :</td>
<td/>
<td style={{padding:"6px 10px",textAlign:"right",color:"#a32d2d",fontWeight:800}}>{rR(totPot)}</td>
</tr>
{ydtRows.length>0&&<>
<tr><td colSpan={4} style={{background:"#F7F7F7",padding:"7px 10px"}}><b style={{textDecoration:"underline"}}>Yang sudah diterima</b></td></tr>
{ydtRows.map((r,i)=><tr key={"y"+i} style={{background:"#F7F7F7"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#666"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,color:"#444",textAlign:"center"}}>{r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:600}}>{Number(r.jumlah||0)>0?rR(r.jumlah)+" -":"Rp -"}</td>
</tr>)}
</>}
</tbody>
</table>
<div style={{padding:"14px 18px",textAlign:"right",fontSize:14,fontWeight:800,borderTop:"1px solid #aaa"}}>
<span style={{marginRight:12,letterSpacing:0.3}}>TOTAL DITERIMA :</span>
<span style={{background:"#FFF35B",padding:"5px 16px",borderRadius:4,border:"1px solid #d4af00",fontSize:16,fontWeight:900}}>{rR(totDiterima)}</span>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",padding:"18px 36px 14px",gap:20,fontSize:13,textAlign:"center"}}>
<div><div style={{color:"#444"}}>Penerima</div><div style={{height:54,margin:"10px 0 2px",display:"flex",alignItems:"center",justifyContent:"center"}}>{slip.ttdPenerima?<img src={slip.ttdPenerima} style={{height:54}} alt="ttd"/>:<div style={{fontSize:10,color:"#bbb",border:"1px dashed #ddd",borderRadius:4,padding:"8px 12px"}}>TTD elektronik</div>}</div><div style={{fontWeight:800,borderTop:"1.5px solid #1a1a1a",paddingTop:3,display:"inline-block",minWidth:150,textDecoration:"underline"}}>{slip.nama}</div></div>
<div><div style={{fontWeight:800}}>{company.nama}</div><div style={{height:54,margin:"10px 0 2px",display:"flex",alignItems:"center",justifyContent:"center"}}>{company.ttdDirektur?<img src={company.ttdDirektur} style={{height:54}} alt="ttd"/>:<div style={{fontSize:10,color:"#bbb",border:"1px dashed #ddd",borderRadius:4,padding:"8px 12px"}}>TTD elektronik</div>}</div><div style={{fontWeight:800,borderTop:"1.5px solid #1a1a1a",paddingTop:3,display:"inline-block",minWidth:150,textDecoration:"underline"}}>{company.direkturNama||"Muhammad Haekal"}</div><div style={{fontSize:12,color:"#444",marginTop:2}}>Direktur</div></div>
</div>
{/* FOOTER — sama dengan Invoice */}
<div style={{background:SNAVY,padding:"12px 20px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
{[["📍","Alamat",[company.alamat]],["📞","No.HP & WA",[company.telepon,"TELP. "+company.telepon2]],["🌐","Email & Website",[company.email,company.website]]].map((x,i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:9}}>
<div style={{width:26,height:26,background:"rgba(255,255,255,.1)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:12}}>{x[0]}</div>
<div><div style={{fontSize:7,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:"rgba(255,255,255,.4)",marginBottom:2}}>{x[1]}</div>{x[2].map((v,j)=><div key={j} style={{fontSize:9,fontWeight:600,color:SWHITE,lineHeight:1.5}}>{v}</div>)}</div>
</div>)}
</div>
</div>
<div style={{maxWidth:680,margin:"12px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<button onClick={()=>doPrint("_slip")} style={{background:SNAVY,color:SWHITE,border:"none",padding:"10px 28px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>🖨️ Cetak / Simpan PDF</button>
<span style={{fontSize:11,color:"#888",alignSelf:"center",fontStyle:"italic"}}>Pilih "Save as PDF" di dialog print</span>
<button onClick={onClose} style={{background:"#566573",color:SWHITE,border:"none",padding:"10px 22px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>✕ Tutup</button>
</div>
</div>;
}

// ─── BERITA ACARA ─────────────────────────────────────────────────────────────
function BeritaAcaraView({ba,company={},onClose}){
if(!ba)return null;
var isTitip=ba.tipe==="titip";var totalQty=(ba.items||[]).reduce((a,it)=>a+Number(it.qty||0),0);
var fileName=safeFileName((ba.konsumenNama||"konsumen")+"_"+(ba.noBA||"BA"));
return <div id="_ba_wrap" style={{position:"fixed",inset:0,background:"#e5e8ec",zIndex:9600,padding:16,overflowY:"auto",fontFamily:"'Times New Roman',serif"}}>
<div id="_ba" style={{maxWidth:680,margin:"0 auto",background:"white",borderRadius:6,padding:"24px 28px",color:"#1a1a1a",boxShadow:"0 6px 30px rgba(0,0,0,.15)",border:"2px solid #1A5276"}}>
<div style={{textAlign:"center",borderBottom:"3px double #1A5276",paddingBottom:14,marginBottom:16}}>
<div style={{display:"flex",justifyContent:"center",marginBottom:6}}><CompanyLogo company={company} h={48}/></div>
<div style={{fontFamily:"Arial",fontWeight:900,fontSize:20,color:"#0A2C5C"}}>{company.nama}</div>
<div style={{fontFamily:"Arial",fontSize:11,color:"#444"}}>{company.alamat}</div>
<div style={{fontFamily:"Arial",fontSize:11,color:"#444"}}>Telp: {company.telepon} | {company.email}</div>
</div>
<div style={{textAlign:"center",marginBottom:16}}><div style={{fontSize:16,fontWeight:700,fontFamily:"Arial",textTransform:"uppercase",color:"#1A5276"}}>BERITA ACARA {isTitip?"PENITIPAN":"PENARIKAN"} TABUNG LPG</div><div style={{fontFamily:"Arial",fontSize:12}}>No. {ba.noBA}</div></div>
<div style={{marginBottom:14,fontFamily:"Arial",fontSize:13,lineHeight:2}}>Yang bertanda tangan di bawah ini menyatakan bahwa pada <b>{fD(ba.tanggal)}</b>, telah terjadi <b style={{color:"#1A5276"}}>{isTitip?"PENITIPAN":"PENARIKAN"} TABUNG LPG</b>:</div>
<div style={{background:"#f8f9ff",border:"1px solid #BDC3C7",borderRadius:6,padding:"12px 16px",marginBottom:14,fontFamily:"Arial",fontSize:13}}>
<div style={{fontWeight:700,color:"#1A5276",marginBottom:6}}>{isTitip?"Penerima Titipan":"Penyerah Tabung"}</div>
<div style={{display:"grid",gridTemplateColumns:"130px 1fr",gap:"4px 8px",lineHeight:1.9}}>
<span>Nama</span><span>: <b>{ba.konsumenNama}</b></span>
<span>No. HP</span><span>: {ba.konsumenTelp||"-"}</span>
<span>Alamat</span><span>: {ba.konsumenAlamat||"-"}</span>
<span>{isTitip?"Diantar oleh":"Ditarik oleh"}</span><span>: <b>{ba.salesNama||"-"}</b></span>
</div>
</div>
<table style={{width:"100%",borderCollapse:"collapse",fontFamily:"Arial",fontSize:13,marginBottom:14}}>
<thead><tr>{["No.","Ukuran","Jumlah"].map((h,i)=><th key={i} style={{background:"#1A5276",color:"#fff",padding:"7px 10px",textAlign:i===2?"center":"left"}}>{h}</th>)}</tr></thead>
<tbody>{(ba.items||[]).map((it,i)=><tr key={i} style={{borderBottom:"1px solid #ddd"}}><td style={{padding:"6px 10px"}}>{i+1}</td><td style={{padding:"6px 10px"}}>Gas LPG {it.ukuran}</td><td style={{padding:"6px 10px",textAlign:"center",fontWeight:700}}>{it.qty}</td></tr>)}<tr style={{background:"#EAF2FF",fontWeight:700}}><td colSpan={2} style={{padding:"8px 10px"}}>TOTAL</td><td style={{padding:"8px 10px",textAlign:"center",fontSize:14}}>{totalQty} Tabung</td></tr></tbody>
</table>
{ba.ket&&<div style={{marginBottom:14,fontFamily:"Arial",fontSize:12,padding:"8px 12px",background:"#FEFDE7",border:"1px solid #F1C40F"}}><b>Ket:</b> {ba.ket}</div>}
<div style={{fontFamily:"Arial",fontSize:12,marginBottom:20}}>Demikian berita acara ini dibuat. <b style={{color:"#1A5276"}}>Tabung merupakan aset milik {company.nama}</b> dan wajib dikembalikan bila tidak lagi berlangganan.</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,fontFamily:"Arial",fontSize:12}}>
{[["Yang Menyerahkan",ba.salesNama,company.ttdKasir],["Yang Menerima",ba.konsumenNama,null]].map((x,i)=><div key={i} style={{textAlign:"center"}}><div style={{fontWeight:600}}>{x[0]}</div><div style={{height:52,display:"flex",alignItems:"center",justifyContent:"center"}}>{x[2]?<img src={x[2]} style={{height:50}} alt="ttd"/>:null}</div><div style={{borderBottom:"1px solid #333",margin:"0 20px"}}/><div style={{marginTop:4,fontWeight:700}}>{x[1]||"____________"}</div></div>)}
</div>
</div>
<div style={{maxWidth:680,margin:"14px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<button onClick={()=>doPrint("_ba")} style={{background:"#1A5276",color:"#fff",border:"none",padding:"10px 28px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / Simpan PDF</button>
<div style={{fontSize:11,color:"#666",alignSelf:"center",fontStyle:"italic"}}>Pilih "Save as PDF" di dialog print</div>
<button onClick={onClose} style={{background:"#566573",color:"#fff",border:"none",padding:"10px 22px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>;
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginScreen({employees,onLogin,themeToggle,theme}){
var C=useTheme();
var[u,setU]=useState("");var[p,setP]=useState("");var[err,setErr]=useState("");
function login(){var emp=(employees||DEF_EMP).find(x=>x.username===u&&x.password===p&&x.aktif);if(emp)onLogin(emp);else setErr("Username atau password salah!");}
return <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:16,position:"relative"}}>
<div style={{position:"absolute",top:16,right:16}}><Btn sm color="gray" onClick={themeToggle}>{theme==="light"?"🌙 Gelap":"☀️ Terang"}</Btn></div>
<div style={{background:C.card,borderRadius:20,padding:"36px 32px 28px",width:"100%",maxWidth:380,border:"1px solid "+C.bdr,boxShadow:"0 20px 80px rgba(0,0,0,.15)"}}>
<div style={{textAlign:"center",marginBottom:26}}><div style={{display:"flex",justifyContent:"center",marginBottom:10}}><LPGLogo size={64}/></div><div style={{fontSize:22,fontWeight:900,color:C.blt}}>PT. HOE TRANG SA</div><div style={{fontSize:13,fontWeight:700,color:C.wht}}>Sistem Manajemen LPG</div><div style={{fontSize:11,color:C.gl2,marginTop:3}}>Distributor Resmi Pertamina</div></div>
<Inp label="Username" value={u} onChange={setU} placeholder="Masukkan username"/>
<Inp label="Password" type="password" value={p} onChange={setP} placeholder="Masukkan password"/>
{err&&<div style={{color:C.rlt,fontSize:12,marginBottom:10,textAlign:"center",padding:"6px 10px",background:C.mode==="dark"?"#3B0A0A":"#FEE2E2",borderRadius:6}}>{err}</div>}
<Btn onClick={login} style={{width:"100%",padding:"12px",fontSize:14,marginBottom:12}}>🔐 Masuk</Btn>
<div style={{background:C.nav,borderRadius:8,padding:10,fontSize:11,color:C.gl2,border:"1px solid "+C.bdr}}><b style={{color:C.wht}}>Contoh:</b> Haekal | DAYAT | MUSLEM | SAIBAN<br/>Password: Sudirman80</div>
</div>
</div>;
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({data,setTab,user}){
var C=useTheme();var td=toDay();
var pHari=(data.penjualan||[]).filter(e=>e.tanggal===td);
var cashIn=pHari.filter(e=>e.bayar==="cash").reduce((a,e)=>a+(e.total||0),0);
var tfIn=pHari.filter(e=>e.bayar==="transfer").reduce((a,e)=>a+(e.total||0),0);
var bonIn=pHari.filter(e=>e.bayar==="bon").reduce((a,e)=>a+(e.total||0),0);
var totalIn=cashIn+tfIn+bonIn;
var margin=pHari.reduce((a,e)=>a+(e.margin||0),0);
var penHari=(data.pengeluaran||[]).filter(e=>e.tanggal===td);
var totalOut=penHari.reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersih=margin-totalOut;
var bonAktif=(data.bon||[]).filter(b=>b.status!=="lunas");
var piutang=bonAktif.reduce((a,b)=>a+b.sisaTagihan,0);
var alerts=bonAktif.filter(b=>b.deadline&&dLeft(b.deadline)!==null&&dLeft(b.deadline)<=3);
var mob=useMobile();
return <div>
<div style={{marginBottom:14}}><div style={{fontSize:20,fontWeight:900,color:C.wht}}>Halo, {user?.nama} 👋</div><div style={{fontSize:12,color:C.gl2,marginTop:2}}>{ROLE_LBL[user?.role]||""} — {fDs(td)}</div></div>
{alerts.length>0&&<div style={{background:C.mode==="dark"?"#3D1A05":"#FFEDD5",border:"1px solid "+C.org,borderRadius:12,padding:"12px 16px",marginBottom:14}}><b style={{color:C.olt,fontSize:13}}>⚠️ {alerts.length} Bon Jatuh Tempo!</b>{alerts.map(b=>{var d=dLeft(b.deadline);return <div key={b.id} style={{fontSize:12,color:C.gltr,marginTop:4}}><b style={{color:C.wht}}>{b.konsumen}</b> — {fR(b.sisaTagihan)} — {d<=0?<Bdg color="red">LEWAT</Bdg>:<Bdg color="orange">{d}h lagi</Bdg>}</div>;})}</div>}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 P&L Hari Ini — {fDs(td)}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8}}>
{[["Omzet",totalIn,C.wht,"📈"],["Laba Kotor",margin,C.blt,"💹"],["Pengeluaran",totalOut,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"9px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2,marginBottom:2}}>{x[3]} {x[0]}</div><div style={{fontSize:14,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
</Card>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
<SC label="Piutang Aktif" value={fR(piutang)} icon="💳" color={C.olt}/>
<SC label="Bon Aktif" value={bonAktif.length+" bon"} icon="📃" color={C.gl2}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:12,marginBottom:14}}>
{SIZES.map(s=>{var isi=(data.stock||{})[s]||0;var asset=(data.totalAsset||{})[s]||0;var titip=getTitipTotal(data.titipList,s);var kosong=getKosong(data,s);return <Card key={s} style={{marginBottom:0}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>📦 LPG {s}</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5}}>{[["Isi",isi,C.glt],["Titip",titip,C.blt],["Kosong",kosong,C.gl2],["Asset",asset,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"5px 6px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:15,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}</div></Card>;})}
</div>
<Card style={{marginBottom:14}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📊 Mutasi Stok Hari Ini — {fDs(td)}</div>
<TabelStokHarian data={data} tgl={td}/>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:10}}>⚡ Akses Cepat</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{[["penjualan","🧾","Penjualan"],["stok","📦","Stok"],["piutang","💳","Piutang"],["absensi","📅","Absensi"],["payroll","💼","Payroll"],["laporan","📊","Laporan"]].map(x=><Btn key={x[0]} onClick={()=>setTab(x[0])} sm>{x[1]} {x[2]}</Btn>)}</div></Card>
</div>;
}

// ─── PENJUALAN v4 (FilterTbl, cetak invoice langsung dari riwayat) ────────────
function PenjualanMod({data,setData,setInv,user,toast}){
var C=useTheme();var mob=useMobile();
var blk={ukuran:"5.5 kg",jenis:"Isi",qty:"",price:""};
var canSelf=PENJUALAN_ROLES.includes(user?.role)&&!["owner","admin","akuntan"].includes(user?.role);
var[f,setF]=useState({tanggal:toDay(),salesId:canSelf?user.id:"",konsumen:"",konsumenId:"",items:[{...blk}],bayar:"cash",bank:"BSI",deadline:"",ket:""});
var[delId,setDelId]=useState(null);
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",bayar:""});
var salesEmp=sortEmp((data.employees||[]).filter(e=>e.aktif&&PENJUALAN_ROLES.includes(e.role)));
var valid=f.items.filter(it=>Number(it.qty)>0&&Number(it.price)>0);
var total=iTotal(valid);var margin=calcMargin(valid,data,f.tanggal);
var kNames=[...new Set([...(data.pelanggan||[]).map(p=>p.nama),...(data.penjualan||[]).map(e=>e.konsumen)].filter(Boolean))];
function onKons(nama){var p=(data.pelanggan||[]).find(x=>x.nama===nama);if(p){setF(pv=>{var newItems=pv.items.map(it=>{var h=(Array.isArray(p.hargaKhusus)?p.hargaKhusus:[]).find(x=>x.ukuran===it.ukuran&&x.jenis===it.jenis);if(h)return{...it,price:String(h.harga)};var het=getHET(data,it.ukuran,it.jenis);return{...it,price:het?String(het):it.price};});return{...pv,konsumen:nama,konsumenId:p.id,items:newItems};});}else setF(pv=>({...pv,konsumen:nama,konsumenId:""}));}
function setProduct(i,ukuran,jenis){setF(p=>{var it=p.items.slice();var newIt={...it[i],ukuran,jenis};var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);if(plg){var h=(plg.hargaKhusus||[]).find(x=>x.ukuran===ukuran&&x.jenis===jenis);if(h){newIt.price=String(h.harga);}else newIt.price=String(getHET(data,ukuran,jenis)||"");}else newIt.price=String(getHET(data,ukuran,jenis)||"");it[i]=newIt;return{...p,items:it};});}
function setItem(i,k,v){setF(p=>{var it=p.items.slice();it[i]={...it[i],[k]:v};return{...p,items:it};});}
function makeInvObj(entry){var emp=(data.employees||[]).find(e=>e.id===entry.salesId);var plg=(data.pelanggan||[]).find(x=>x.id===entry.konsumenId);return{noInv:entry.noInv,tanggal:entry.tanggal,konsumen:entry.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",salesNama:emp?.nama||"",items:(entry.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price)})),total:entry.total,metodeBayar:entry.bayar==="bon"?"BON":entry.bayar==="transfer"?"Transfer "+(entry.bank||""):"Cash",isBon:entry.bayar==="bon",catatan:entry.ket||""};}
function doSave(withPrint){
if(!valid.length||!f.konsumen)return;
var ns={...data.stock};
var nk={...(data.stokKosong||{})};
var na={...(data.totalAsset||{})};
var stokLogs=[];
valid.forEach(it=>{
  var q=Number(it.qty||0);var uk=it.ukuran;
  ns[uk]=Math.max(0,(ns[uk]||0)-q);
  if(it.jenis==="Tabung+Isi"){
    nk[uk]=Math.max(0,(nk[uk]||0)-q);
    na[uk]=Math.max(0,(na[uk]||0)-q);
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Tbg+Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  } else {
    nk[uk]=(nk[uk]||0)+q;
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  }
});
var invInfo=nextInvNo(data,f.tanggal);
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.inv)newCounters.inv={};newCounters.inv[invInfo.key]=invInfo.n;
var entry={id:uid(),noInv:invInfo.no,tanggal:f.tanggal,waktu:new Date().toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"}),salesId:f.salesId,konsumen:f.konsumen,konsumenId:f.konsumenId,items:valid.map(it=>({...it,qty:Number(it.qty),price:Number(it.price)})),total,margin,bayar:f.bayar,bank:f.bank,deadline:f.deadline,ket:f.ket};
var nb=(data.bon||[]).slice();
if(f.bayar==="bon")nb.unshift({id:uid(),noInv:invInfo.no,tanggal:f.tanggal,konsumen:f.konsumen,konsumenId:f.konsumenId,salesId:f.salesId,items:valid,total,sisaTagihan:total,deadline:f.deadline,status:"belum",pembayaran:[],ket:f.ket,bank:f.bank});
setData(d=>({...d,penjualan:[entry,...(d.penjualan||[])],stock:ns,stokKosong:nk,totalAsset:na,bon:nb,counters:newCounters,stockLog:[...stokLogs,...(d.stockLog||[])].slice(0,500)}));
if(withPrint)setInv(makeInvObj(entry));
setF(p=>({...p,konsumen:"",konsumenId:"",items:[{...blk}],ket:"",deadline:""}));
toast("✓ Tersimpan! No: "+invInfo.no);
}
// Riwayat: flatten transactions, apply bar filter
var rows=useMemo(()=>{
return(data.penjualan||[]).filter(p=>{
if(barFilter.from&&p.tanggal<barFilter.from)return false;
if(barFilter.to&&p.tanggal>barFilter.to)return false;
if(barFilter.salesId&&p.salesId!==barFilter.salesId)return false;
if(barFilter.konsumen&&!p.konsumen.toLowerCase().includes(barFilter.konsumen.toLowerCase()))return false;
if(barFilter.bayar&&p.bayar!==barFilter.bayar)return false;
return true;
}).map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var detailStr=(p.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ");return{...p,salesNama:emp?.nama||"-",detailStr};});
},[data.penjualan,data.employees,barFilter]);
var SZ_CLR={"5.5 kg":[C.grn,"#0A2E14"],"12 kg":[C.blt,"#0A2040"],"50 kg":[C.olt,"#3D200A"]};
var cols=[
{key:"tanggal",label:"Tgl",width:72,render:r=><div><div style={{fontSize:12,fontWeight:600,color:C.wht,whiteSpace:"nowrap"}}>{fDs(r.tanggal)}</div><div style={{fontSize:10,color:C.gl2,marginTop:2}}>{r.waktu||""}</div></div>,sortVal:r=>r.tanggal,filterable:true},
{key:"noInv",label:"No.Invoice",width:69,render:r=>{var parts=(r.noInv||"-").split("/").reduce((a,p,i)=>i===0?[p]:i===1?[a[0]+"/"+p]:i===2?[...a,p+"/"]:[...a.slice(0,-1),a[a.length-1]+p],[]);var half=Math.ceil((r.noInv||"").split("/").length/2);var top=(r.noInv||"").split("/").slice(0,2).join("/");var bot=(r.noInv||"").split("/").slice(2).join("/");return <div style={{fontSize:11,color:C.blt,fontWeight:700,lineHeight:1.4}}><div>{top}/</div><div>{bot}</div></div>;},filterable:true},
{key:"konsumen",label:"Konsumen",render:r=><b style={{color:C.wht,display:"block",minWidth:120}}>{r.konsumen}</b>,filterable:true,width:127},
{key:"salesNama",label:"Sales",render:r=><span style={{fontSize:11}}>{r.salesNama}</span>,filterable:true,width:69},
{key:"detailStr",label:"Produk & Qty",render:r=>{var items=r.items||[];var show=items.slice(0,3);var more=items.length-3;return <div style={{display:"flex",flexDirection:"column",gap:3,width:125,maxWidth:125}}>{show.map((it,i)=>{var clr=SZ_CLR[it.ukuran]||[C.gl2,C.bg];return <div key={i} style={{display:"flex",alignItems:"center",gap:4,paddingBottom:i<show.length-1?2:0,borderBottom:i<show.length-1?"1px dashed "+C.bdr:"none"}}><span style={{background:clr[1],border:"1px solid "+clr[0],borderRadius:3,padding:"1px 5px",fontSize:9,fontWeight:700,color:clr[0],whiteSpace:"nowrap",flexShrink:0,minWidth:42,textAlign:"center"}}>{it.ukuran}</span><span style={{fontSize:10.5,color:C.gl2,whiteSpace:"nowrap",flexShrink:0}}>{it.jenis==="Tabung+Isi"?"Tbg+Isi":"Refill"}</span><b style={{fontSize:12,fontWeight:800,color:C.wht,whiteSpace:"nowrap",marginLeft:"auto"}}>{it.qty}</b></div>;})}{ more>0&&<div style={{fontSize:10,color:C.gry,fontStyle:"italic"}}>+{more} item lagi</div>}</div>;},filterable:true,sortable:false},
{key:"total",label:"Total",render:r=><b style={{color:C.wht,whiteSpace:"nowrap"}}>{fR(r.total)}</b>,filterable:false,sortVal:r=>r.total,width:98},
{key:"bayar",label:"Bayar",render:r=>r.bayar==="bon"?<Bdg color="red">BON</Bdg>:r.bayar==="transfer"?<Bdg color="blue">TF</Bdg>:<Bdg color="green">Cash</Bdg>,filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"}],width:47},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,width:63,render:r=><div style={{display:"flex",gap:4}}><button onClick={()=>setInv(makeInvObj(r))} title="Cetak Invoice" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button><button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button></div>},
];
return <div>
<STitle icon="🧾" children="Input Penjualan"/>
<Card>
<div style={{display:"grid",gridTemplateColumns:mob?"1fr 1fr":"repeat(auto-fill,minmax(155px,1fr))",gap:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/>
<Sel label="Penjual" value={f.salesId} onChange={v=>setF(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...salesEmp.map(e=>({v:e.id,l:e.nama+" ("+e.posisi+")"}))]} style={{gridColumn:mob?"1/-1":"auto"}}/>
<div style={{gridColumn:mob?"1/-1":"auto"}}><AutoInp label="Nama Konsumen" value={f.konsumen} onChange={v=>setF(p=>({...p,konsumen:v}))} options={kNames} placeholder="Ketik nama..." onSelect={onKons}/></div>
</div>
{f.konsumenId&&(()=>{var plg=(data.pelanggan||[]).find(x=>x.id===f.konsumenId);return plg&&(Array.isArray(plg.hargaKhusus)?plg.hargaKhusus:[]).length>0?<div style={{padding:"6px 10px",background:C.mode==="dark"?"#0A1A2A":"#DBEAFE",borderRadius:6,fontSize:11,color:C.blt,marginBottom:10}}>💲 Pelanggan memiliki <b>{plg.hargaKhusus.length}</b> harga khusus — otomatis terpasang</div>:null;})()}
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
<div style={{display:"grid",gridTemplateColumns:mob?"85px 100px 55px 95px 28px":"95px 110px 60px 110px 90px 28px",background:C.nav,padding:"6px 10px",fontSize:11,color:C.gl2,fontWeight:700,gap:5}}><span>Ukuran</span><span>Jenis</span><span>Qty</span><span>Harga</span>{!mob&&<span>Subtotal</span>}<span/></div>
{f.items.map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:mob?"85px 100px 55px 95px 28px":"95px 110px 60px 110px 90px 28px",padding:"5px 10px",borderTop:"1px solid "+C.bdr,alignItems:"center",gap:5}}>
<select value={it.ukuran} onChange={e=>setProduct(i,e.target.value,it.jenis)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none"}}>{SIZES.map(s=><option key={s}>{s}</option>)}</select>
<select value={it.jenis} onChange={e=>setProduct(i,it.ukuran,e.target.value)} style={{background:C.nav,border:"1px solid "+(it.jenis==="Tabung+Isi"?C.glt:C.bdr),borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none"}}>{JENIS.map(j=><option key={j}>{j}</option>)}</select>
<input type="number" value={it.qty} placeholder="0" onChange={e=>setItem(i,"qty",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<input type="number" value={it.price} onChange={e=>setItem(i,"price",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
{!mob&&<span style={{color:C.glt,fontWeight:700,fontSize:12}}>{it.qty&&it.price?fR(Number(it.qty)*Number(it.price)):"-"}</span>}
<button onClick={()=>setF(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} disabled={f.items.length<=1} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13,padding:"2px 5px",opacity:f.items.length<=1?0.3:1}}>✕</button>
</div>)}
<div style={{padding:"7px 10px",background:C.nav,borderTop:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
<Btn sm color="blue" onClick={()=>setF(p=>({...p,items:[...p.items,{...blk}]}))}>+ Item</Btn>
<span style={{fontSize:13,color:C.gl2}}>Margin: <b style={{color:C.glt}}>{fR(margin)}</b> | Total: <b style={{color:C.wht,fontSize:14}}>{fR(total)}</b></span>
</div>
</div>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu],["bon","📃 BON",C.rdk]].map(x=><button key={x[0]} onClick={()=>setF(p=>({...p,bayar:x[0]}))} style={{background:f.bayar===x[0]?x[2]:C.nav,color:f.bayar===x[0]?"white":C.wht,border:"1px solid "+(f.bayar===x[0]?x[2]:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:13,cursor:"pointer"}}>{x[1]}</button>)}</div>
{f.bayar==="transfer"&&<div style={{display:"flex",gap:8,marginBottom:10}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,bank:b}))} style={{background:f.bank===b?C.blu:C.nav,color:f.bank===b?"white":C.wht,border:"2px solid "+(f.bank===b?C.blt:C.bdr),borderRadius:8,padding:"6px 16px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
{f.bayar==="bon"&&<Inp label="Deadline" type="date" value={f.deadline} onChange={v=>setF(p=>({...p,deadline:v}))} style={{maxWidth:220}}/>}
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))} placeholder="Catatan opsional"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
<Btn onClick={()=>doSave(false)} color="green" dis={!valid.length||!f.konsumen}>💾 Simpan</Btn>
<Btn onClick={()=>doSave(true)} color="blue" dis={!valid.length||!f.konsumen}>🖨️ Simpan & Cetak Invoice</Btn>
</div>
</Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Penjualan</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Sales" value={barFilter.salesId} onChange={v=>setBarFilter(p=>({...p,salesId:v}))} opts={[{v:"",l:"Semua"},...salesEmp.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
<Inp label="Konsumen" value={barFilter.konsumen} onChange={v=>setBarFilter(p=>({...p,konsumen:v}))} placeholder="Cari..." style={{marginBottom:0}}/>
<Sel label="Bayar" value={barFilter.bayar} onChange={v=>setBarFilter(p=>({...p,bayar:v}))} opts={[{v:"",l:"Semua"},{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"}]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.salesId||barFilter.konsumen||barFilter.bayar)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",salesId:"",konsumen:"",bayar:""})}>✕ Reset Bar Filter</Btn></div>}
</div>
<FilterTbl columns={cols} data={rows} empty="Belum ada penjualan" maxRows={150}/>
</Card>
{delId&&<ConfirmDel msg={"Hapus penjualan \""+delId.konsumen+"\"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,penjualan:(d.penjualan||[]).filter(x=>x.id!==delId.id)}));setDelId(null);toast("✓ Dihapus");}}/>}
</div>;
}

// === AKHIR BAGIAN 2 ===
// === BAGIAN 3 DARI 3 ===

// ─── STOK ─────────────────────────────────────────────────────────────────────
function StokMod({data,setData,user,toast}){
var C=useTheme();var[tab,setTab]=useState("rekap");var[ba,setBa]=useState(null);
function RekapTab(){
var[initMode,setInitMode]=useState(false);var[asF,setAsF]=useState({"5.5 kg":"","12 kg":"","50 kg":""});
return <div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:12,marginBottom:14}}>
{SIZES.map(s=>{var isi=(data.stock||{})[s]||0;var asset=(data.totalAsset||{})[s]||0;var titip=getTitipTotal(data.titipList,s);var kosong=getKosong(data,s);var pI=asset>0?Math.round(isi/asset*100):0;var pT=asset>0?Math.round(titip/asset*100):0;var pK=asset>0?Math.round(kosong/asset*100):0;return <Card key={s} style={{marginBottom:0}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><div style={{fontWeight:800,color:C.wht,fontSize:15}}>📦 LPG {s}</div><div style={{textAlign:"right"}}><div style={{fontSize:10,color:C.gl2}}>Total Asset</div><div style={{fontSize:18,fontWeight:900,color:C.olt}}>{asset}</div></div></div><div style={{height:8,borderRadius:4,background:C.bdr,display:"flex",overflow:"hidden",marginBottom:10}}><div style={{width:pI+"%",background:C.glt}}/><div style={{width:pT+"%",background:C.blt}}/><div style={{width:pK+"%",background:C.gl2}}/></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>{[["🟢 Isi",isi,C.glt],["🔵 Titip",titip,C.blt],["⚫ Kosong",kosong,C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"7px 6px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:20,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}</div></Card>;})}
</div>
<Card><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><div style={{fontWeight:700,color:C.gl2,fontSize:13}}>🏭 Setup Total Asset</div><Btn sm color={initMode?"gray":"blue"} onClick={()=>setInitMode(!initMode)}>{initMode?"Batal":"⚙️ Set"}</Btn></div>
{initMode&&<div><div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Input stok fisik awal per ukuran:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10}}>
{SIZES.map(s=>[
<Inp key={"isi_"+s} label={"Isi "+s} type="number" value={asF["isi_"+s]||""} placeholder={String((data.stock||{})[s]||0)} onChange={v=>setAsF(p=>({...p,["isi_"+s]:v}))}/>,
<Inp key={"kos_"+s} label={"Kosong "+s} type="number" value={asF["kos_"+s]||""} placeholder={String(getKosong(data,s))} onChange={v=>setAsF(p=>({...p,["kos_"+s]:v}))}/>
])}
</div>
<Btn color="green" onClick={()=>{
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
SIZES.forEach(s=>{
if(asF["isi_"+s]!=="")ns[s]=Number(asF["isi_"+s]);
if(asF["kos_"+s]!=="")nk[s]=Number(asF["kos_"+s]);
});
var na={};SIZES.forEach(s=>{na[s]=(nk[s]||0)+getTitipTotal(data.titipList,s);});
setData(d=>({...d,stock:ns,stokKosong:nk,totalAsset:na}));
setInitMode(false);toast("✓ Stok awal disimpan!");
}}>💾 Simpan Stok Awal</Btn></div>}
</Card>
</div>;
}
function MutasiTab(){
var[f,setF]=useState({ukuran:"12 kg",jenis:"return_isi",qty:"",ket:"",tanggal:toDay()});var[delId,setDelId]=useState(null);
function save(){
if(!f.qty||Number(f.qty)<=0)return;
var qty=Number(f.qty);var s=f.ukuran;
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalAsset||{})};
var jDesc="";
if(f.jenis==="return_isi"){ns[s]=(ns[s]||0)+qty;jDesc="↩️ Return (+Isi)";}
else if(f.jenis==="pancung"){ns[s]=(ns[s]||0)+qty;jDesc="✂️ Pancung (+Isi)";}
else if(f.jenis==="beli_tbg"){nk[s]=(nk[s]||0)+qty;na[s]=(na[s]||0)+qty;jDesc="🛒 Beli Tabung (+Tbg)";}
else if(f.jenis==="rusak"){ns[s]=Math.max(0,(ns[s]||0)-qty);jDesc="💥 Rusak/Bocor (-Isi)";}
var log={id:uid(),tanggal:f.tanggal,ukuran:s,jenis:jDesc,qty,ket:f.ket,user:user?.nama||"",sumber:"Manual"};
setData(d=>({...d,stock:ns,stokKosong:nk,totalAsset:na,stockLog:[log,...(d.stockLog||[])].slice(0,500)}));
setF(p=>({...p,qty:"",ket:""}));
toast("✓ Mutasi dicatat!");
}
return <div><Card><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10}}><Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/><Sel label="Ukuran" value={f.ukuran} onChange={v=>setF(p=>({...p,ukuran:v}))} opts={SIZES}/><Sel label="Jenis" value={f.jenis} onChange={v=>setF(p=>({...p,jenis:v}))} opts={[{v:"return_isi",l:"↩️ Return (+Isi)"},{v:"pancung",l:"✂️ Pancung (+Isi)"},{v:"beli_tbg",l:"🛒 Beli Tabung dr Konsumen (+Tbg)"},{v:"rusak",l:"💥 Rusak/Bocor (-Isi)"}]}/><Inp label="Qty" type="number" value={f.qty} onChange={v=>setF(p=>({...p,qty:v}))}/><Inp label="Ket" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/></div><Btn color="green" onClick={save} dis={!f.qty}>💾 Simpan Mutasi</Btn></Card><Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Log Mutasi</div><RTbl headers={["Tgl","Ukuran","Jenis","Qty","Ket","Aksi"]} rows={(data.stockLog||[]).slice(0,50).map(l=>[fDs(l.tanggal),l.ukuran,l.jenis,l.qty,l.ket||"-",<ActBtns onDel={()=>setDelId(l)}/>])}/></Card>{delId&&<ConfirmDel msg="Hapus log?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,stockLog:(d.stockLog||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}</div>;
}
function OpnameTab(){
var C=useTheme();
var[opDate,setOpDate]=useState(toDay());
var[opF,setOpF]=useState(()=>{var o={};SIZES.forEach(s=>{o["isi_"+s]="";o["kos_"+s]="";});return o;});
var[opResult,setOpResult]=useState(null);

function hitungSelisih(){
  var results=SIZES.map(s=>{
    var sysIsi=(data.stock||{})[s]||0;
    var sysKos=getKosong(data,s);
    var fisIsi=opF["isi_"+s]!==""?Number(opF["isi_"+s]):null;
    var fisKos=opF["kos_"+s]!==""?Number(opF["kos_"+s]):null;
    var selIsi=fisIsi!==null?fisIsi-sysIsi:null;
    var selKos=fisKos!==null?fisKos-sysKos:null;
    return{ukuran:s,sysIsi,sysKos,fisIsi,fisKos,selIsi,selKos};
  });
  setOpResult({tanggal:opDate,results});
}

function doAdjust(s,type,sel){
  var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
  var jDesc="";
  if(type==="isi"){ns[s]=Math.max(0,(ns[s]||0)+sel);jDesc="Adj Opname Isi ("+fDs(opDate)+")";}
  else{nk[s]=Math.max(0,(nk[s]||0)+sel);jDesc="Adj Opname Kosong ("+fDs(opDate)+")";}
  var log={id:uid(),tanggal:opDate,ukuran:s,jenis:jDesc,qty:Math.abs(sel),ket:"Selisih Opname Checker",user:user?.nama||"",sumber:"Opname"};
  setData(d=>({...d,stock:ns,stokKosong:nk,stockLog:[log,...(d.stockLog||[])].slice(0,500)}));
  // Update hasil opname display
  setOpResult(prev=>prev?{...prev,results:prev.results.map(r=>{
    if(r.ukuran!==s)return r;
    var newSysIsi=type==="isi"?Math.max(0,r.sysIsi+sel):r.sysIsi;
    var newSysKos=type==="kos"?Math.max(0,r.sysKos+sel):r.sysKos;
    return{...r,sysIsi:newSysIsi,sysKos:newSysKos,selIsi:r.fisIsi!==null?r.fisIsi-newSysIsi:null,selKos:r.fisKos!==null?r.fisKos-newSysKos:null};
  })}:null);
  toast("✓ Stok disesuaikan!");
}

return <div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🔍 Opname Harian — Input Stok Fisik</div>
<Inp label="Tanggal Opname" type="date" value={opDate} onChange={setOpDate} style={{maxWidth:200,marginBottom:12}}/>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:12}}>
{SIZES.map(s=>[
  <Inp key={"oi"+s} label={"Fisik Isi "+s+" (Sistem: "+(((data.stock||{})[s]||0))+")"} type="number" value={opF["isi_"+s]} onChange={v=>setOpF(p=>({...p,["isi_"+s]:v}))} placeholder="kosongkan jika tidak opname"/>,
  <Inp key={"ok"+s} label={"Fisik Kosong "+s+" (Sistem: "+getKosong(data,s)+")"} type="number" value={opF["kos_"+s]} onChange={v=>setOpF(p=>({...p,["kos_"+s]:v}))} placeholder="kosongkan jika tidak opname"/>
])}
</div>
<Btn color="blue" onClick={hitungSelisih}>🔍 Hitung Selisih</Btn>
</Card>

{opResult&&<Card style={{border:"1px solid #2980B9"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>📊 Hasil Opname — {fDs(opResult.tanggal)}</div>
{opResult.results.map(r=>{
  var adaSelisih=(r.selIsi!==null&&r.selIsi!==0)||(r.selKos!==null&&r.selKos!==0);
  return <div key={r.ukuran} style={{marginBottom:10,padding:12,background:C.nav,borderRadius:8,border:"1px solid "+(adaSelisih?C.rlt:C.glt)}}>
  <div style={{fontWeight:800,color:C.wht,marginBottom:8}}>📦 {r.ukuran}</div>
  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:adaSelisih?8:4}}>
  {[["Isi Sistem",r.sysIsi,C.glt],["Isi Fisik",r.fisIsi!==null?r.fisIsi:"—",C.glt],["Kosong Sistem",r.sysKos,C.gl2],["Kosong Fisik",r.fisKos!==null?r.fisKos:"—",C.gl2]].map(x=><div key={x[0]} style={{background:C.bg,borderRadius:6,padding:"6px 8px",textAlign:"center"}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:16,fontWeight:800,color:x[2]}}>{x[1]}</div></div>)}
  </div>
  {adaSelisih?<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
    {r.selIsi!==null&&r.selIsi!==0&&<div style={{flex:1,minWidth:160,padding:"8px 12px",background:C.rdk,borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
    <span style={{color:"white",fontSize:12,fontWeight:700}}>💨 Isi: {r.selIsi>0?"+":""}{r.selIsi} tab</span>
    <button onClick={()=>doAdjust(r.ukuran,"isi",r.selIsi)} style={{background:"white",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,color:"#DC2626",cursor:"pointer"}}>✓ Adjust</button>
    </div>}
    {r.selKos!==null&&r.selKos!==0&&<div style={{flex:1,minWidth:160,padding:"8px 12px",background:C.rdk,borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
    <span style={{color:"white",fontSize:12,fontWeight:700}}>📦 Kosong: {r.selKos>0?"+":""}{r.selKos} tab</span>
    <button onClick={()=>doAdjust(r.ukuran,"kos",r.selKos)} style={{background:"white",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,color:"#DC2626",cursor:"pointer"}}>✓ Adjust</button>
    </div>}
  </div>:<div style={{color:C.glt,fontSize:12,fontWeight:700}}>✅ Stok sesuai — tidak ada selisih</div>}
  </div>;
})}
</Card>}

<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Log Opname</div>
<RTbl headers={["Tgl","Ukuran","Keterangan","Qty","User"]} rows={(data.stockLog||[]).filter(l=>l.sumber==="Opname").slice(0,30).map(l=>[fDs(l.tanggal),l.ukuran,l.jenis,l.qty,l.user||"-"])}/>
</Card>
</div>;
}
function TitipTab(){
var blkI={ukuran:"12 kg",qty:""};
var[f,setF]=useState({tanggal:toDay(),tipe:"titip",konsumenNama:"",konsumenTelp:"",konsumenAlamat:"",salesId:"",items:[{...blkI}],ket:""});
var[delId,setDelId]=useState(null);
var[tf,setTf]=useState({from:"",to:"",tipe:"",ukuran:""});
var kNames=[...new Set([...(data.pelanggan||[]).filter(p=>PLG_TITIP_KAT.includes(p.kategori)).map(p=>p.nama),...(data.titipList||[]).map(t=>t.konsumenNama)].filter(Boolean))];
function onKons(nama){var p=(data.pelanggan||[]).find(x=>x.nama===nama);if(p)setF(pv=>({...pv,konsumenNama:nama,konsumenTelp:p.telepon||"",konsumenAlamat:p.alamat||""}));}
function setItem(i,k,v){setF(p=>{var it=p.items.slice();it[i]={...it[i],[k]:v};return{...p,items:it};});}
var validItems=f.items.filter(it=>Number(it.qty)>0);
function save(cetak){if(!f.konsumenNama||!validItems.length)return;var emp=(data.employees||[]).find(e=>e.id===f.salesId);var noBA="BA-"+Date.now().toString(36).toUpperCase().slice(-6);var rec={id:uid(),noBA,tanggal:f.tanggal,tipe:f.tipe,konsumenNama:f.konsumenNama,konsumenTelp:f.konsumenTelp,konsumenAlamat:f.konsumenAlamat,salesId:f.salesId,salesNama:emp?.nama||"",items:validItems.map(it=>({ukuran:it.ukuran,qty:Number(it.qty)})),ket:f.ket};setData(d=>({...d,titipList:[rec,...(d.titipList||[])]}));if(cetak)setBa(rec);else toast("✓ Dicatat!");setF(p=>({...p,konsumenNama:"",konsumenTelp:"",konsumenAlamat:"",items:[{...blkI}],ket:""}));}
var balMap=getKonsumenTitipBal(data.titipList);var aktifK=Object.keys(balMap).filter(k=>SIZES.some(s=>(balMap[k][s]||0)>0));
return <div>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}><SC label="Konsumen Titip" value={aktifK.length} icon="🏪" color={C.blt}/>{SIZES.map(s=><SC key={s} label={"Titip "+s} value={getTitipTotal(data.titipList,s)+" tab"} icon="📦" color={C.blt}/>)}</div>
<Card>
<div style={{display:"flex",gap:6,marginBottom:12}}>{[["titip","📦 Titip ke Konsumen",C.grn],["tarik","↩️ Tarik dari Konsumen",C.rdk],["titip_luar","🏭 Titipan Pihak Lain Masuk",C.blt],["tarik_luar","📤 Titipan Pihak Lain Keluar","#6B7280"]].map(x=><button key={x[0]} onClick={()=>setF(p=>({...p,tipe:x[0]}))} style={{background:f.tipe===x[0]?x[2]:C.nav,color:f.tipe===x[0]?"white":C.wht,border:"1px solid "+(f.tipe===x[0]?x[2]:C.bdr),borderRadius:8,padding:"8px 16px",fontWeight:700,fontSize:13,cursor:"pointer",flex:1}}>{x[1]}</button>)}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/>
<Sel label="Sales" value={f.salesId} onChange={v=>setF(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).filter(e=>e.aktif).map(e=>({v:e.id,l:e.nama}))]}/>
<div style={{gridColumn:"1/-1"}}><AutoInp label="Konsumen" value={f.konsumenNama} onChange={v=>setF(p=>({...p,konsumenNama:v}))} options={kNames} placeholder="Ketik..." onSelect={onKons}/></div>
<Inp label="No. HP" value={f.konsumenTelp} onChange={v=>setF(p=>({...p,konsumenTelp:v}))}/>
<Inp label="Alamat" value={f.konsumenAlamat} onChange={v=>setF(p=>({...p,konsumenAlamat:v}))}/>
</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 36px",background:C.nav,padding:"6px 10px",fontSize:11,color:C.gl2,fontWeight:700,gap:6}}><span>Ukuran</span><span>Qty</span><span/></div>
{f.items.map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 36px",padding:"5px 10px",borderTop:"1px solid "+C.bdr,gap:6}}>
<select value={it.ukuran} onChange={e=>setItem(i,"ukuran",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 6px",color:C.wht,fontSize:12,outline:"none"}}>{SIZES.map(s=><option key={s}>{s}</option>)}</select>
<input type="number" value={it.qty} placeholder="0" onChange={e=>setItem(i,"qty",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 6px",color:C.wht,fontSize:12,outline:"none"}}/>
<button onClick={()=>setF(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} disabled={f.items.length<=1} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13}}>✕</button>
</div>)}
<div style={{padding:"6px 10px",background:C.nav,borderTop:"1px solid "+C.bdr}}><Btn sm color="blue" onClick={()=>setF(p=>({...p,items:[...p.items,{...blkI}]}))}>+ Item</Btn></div>
</div>
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn color="green" onClick={()=>save(false)} dis={!f.konsumenNama||!validItems.length}>💾 Simpan</Btn><Btn color="blue" onClick={()=>save(true)} dis={!f.konsumenNama||!validItems.length}>🖨️ Simpan & Cetak BA</Btn></div>
</Card>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Riwayat Titip/Tarik</div>
</div>
{(()=>{
var titipFilt=(data.titipList||[]).filter(t=>{
if(tf.from&&(t.tanggal||"")<tf.from)return false;
if(tf.to&&(t.tanggal||"")>tf.to)return false;
if(tf.tipe&&t.tipe!==tf.tipe)return false;
if(tf.ukuran){var items=t.items&&t.items.length>0?t.items:[{ukuran:t.ukuran}];if(!items.some(i=>i.ukuran===tf.ukuran))return false;}
return true;
});
return <>
<div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10,background:C.nav,padding:8,borderRadius:8,border:"1px solid "+C.bdr}}>
<Inp label="Dari" type="date" value={tf.from} onChange={v=>setTf(p=>({...p,from:v}))} style={{marginBottom:0,minWidth:130}}/>
<Inp label="Sampai" type="date" value={tf.to} onChange={v=>setTf(p=>({...p,to:v}))} style={{marginBottom:0,minWidth:130}}/>
<Sel label="Tipe" value={tf.tipe} onChange={v=>setTf(p=>({...p,tipe:v}))} opts={[{v:"",l:"Semua"},{v:"titip",l:"Titip"},{v:"tarik",l:"Tarik"}]}/>
<Sel label="Ukuran" value={tf.ukuran} onChange={v=>setTf(p=>({...p,ukuran:v}))} opts={[{v:"",l:"Semua Ukuran"},{v:"5.5 kg",l:"5,5 kg"},{v:"12 kg",l:"12 kg"},{v:"50 kg",l:"50 kg"}]}/>
{(tf.from||tf.to||tf.tipe||tf.ukuran)&&<Btn sm color="gray" onClick={()=>setTf({from:"",to:"",tipe:"",ukuran:""})}>✕ Reset</Btn>}
</div>
<RTbl headers={["Tgl","Tipe","Konsumen","Sales","Ukuran","Qty","Aksi"]} rows={titipFilt.slice(0,100).map(t=>{
var items=t.items&&t.items.length>0?t.items:[{ukuran:t.ukuran,qty:t.qty}];
var validItems=items.filter(i=>i.ukuran&&i.qty);
var m=t.tipe==="titip"?1:-1;
return[
fDs(t.tanggal),
<Bdg color={t.tipe==="titip"?"green":t.tipe==="tarik"?"red":"blue"}>{t.tipe}</Bdg>,
<b style={{color:C.wht}}>{t.konsumenNama}</b>,
t.salesPengantar||t.salesNama||"-",
<div style={{display:"flex",flexDirection:"column",gap:2}}>{validItems.map((it,i)=><Bdg key={i} color={it.ukuran==="12 kg"?"blue":it.ukuran==="5.5 kg"?"green":"orange"}>{it.ukuran}</Bdg>)}</div>,
<div style={{display:"flex",flexDirection:"column",gap:2}}>{validItems.map((it,i)=><b key={i} style={{color:m>0?C.glt:C.rlt}}>{m>0?"+":"-"}{it.qty}</b>)}</div>,
<div style={{display:"flex",gap:4}}><button onClick={()=>setBa(t)} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button><ActBtns onDel={()=>setDelId(t)}/></div>
];})}/>
</>;
})()}
</Card>
{delId&&<ConfirmDel msg="Hapus?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,titipList:(d.titipList||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}
return <div>
<STitle icon="📦" children="Stok & Mutasi"/>
<div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>{[["rekap","📊 Rekap"],["mutasi","✏️ Mutasi Manual"],["opname","🔍 Opname"],["titip","🏪 Titip/Tarik"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="rekap"&&<RekapTab/>}
{tab==="mutasi"&&<MutasiTab/>}
{tab==="opname"&&<OpnameTab/>}
{tab==="titip"&&<TitipTab/>}
{ba&&<BeritaAcaraView ba={ba} company={data.company} onClose={()=>setBa(null)}/>}
</div>;
}

// ─── DO v4 (HPP/unit + Total HPP, push ke modalHistory) ───────────────────────
function DOMod({data,setData,user,toast}){
var C=useTheme();
var[f,setF]=useState(()=>{var hpp=getModal(data,"12 kg","Isi",toDay());return{tanggal:toDay(),trip:"Trip 1",sppbe:"SPPBE KCR",ukuran:"12 kg",qty:"",hppUnit:String(hpp||""),ket:"",dibuatOleh:user?.id||""};});
var[delId,setDelId]=useState(null);
function onUkuran(v){var hpp=getModal(data,v,"Isi",f.tanggal);setF(p=>({...p,ukuran:v,hppUnit:String(hpp||"")}));}
function onTanggal(v){var hpp=getModal(data,f.ukuran,"Isi",v);setF(p=>({...p,tanggal:v,hppUnit:String(hpp||"")}));}
var totalHPP=Number(f.qty||0)*Number(f.hppUnit||0);
function save(){
if(!f.qty||!f.trip)return;
var qty=Number(f.qty);var hpp=Number(f.hppUnit||0);
var empDO=(data.employees||[]).find(e=>e.id===f.dibuatOleh);
var doRec={id:uid(),tanggal:f.tanggal,trip:f.trip,noDO:f.trip,sppbe:f.sppbe,ukuran:f.ukuran,qty,hppUnit:hpp,totalHPP:qty*hpp,ket:f.ket,dibuatOleh:empDO?.nama||user?.nama||"",status:"gantung"};
var newModalHistory=hpp>0?[{id:uid(),tanggal:f.tanggal,ukuran:f.ukuran,jenis:"Isi",harga:hpp,sumber:"DO "+f.trip},...(data.modalHistory||[])]:(data.modalHistory||[]);
// Stok TIDAK langsung masuk — harus klik Diterima
setData(d=>({...d,doList:[doRec,...(d.doList||[])],modalHistory:newModalHistory}));
setF(p=>({...p,qty:"",ket:""}));
toast("✓ DO dibuat! Status: Gantung. Klik '✅ Diterima' setelah barang tiba di gudang.");
}

// Terima DO → stok masuk
function terimaDO(d_rec){
var qty=Number(d_rec.qty||0);var uk=d_rec.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:d_rec.tanggal,ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+d_rec.trip+" - "+d_rec.sppbe+" (Diterima)",sumber:"DO",user:user?.nama||""};
setData(d=>({...d,
  doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"diterima",tglDiterima:toDay()}:x),
  stock:ns,
  stockLog:[log,...(d.stockLog||[])].slice(0,500)
}));
toast("✅ DO diterima! Stok "+uk+" +"+qty+" tabung masuk gudang.");
}

// Tandai Sangkut
function sangkutDO(d_rec){
setData(d=>({...d,doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"sangkut"}:x)}));
toast("⚠️ DO ditandai Sangkut.");
}

// Release Sangkut → stok masuk
function releaseDO(d_rec){
var qty=Number(d_rec.qty||0);var uk=d_rec.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:toDay(),ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+d_rec.trip+" Release Sangkut",sumber:"DO",user:user?.nama||""};
setData(d=>({...d,
  doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"diterima",tglDiterima:toDay()}:x),
  stock:ns,
  stockLog:[log,...(d.stockLog||[])].slice(0,500)
}));
toast("✅ DO di-release! Stok "+uk+" +"+qty+" tabung masuk.");
}
return <div>
<div style={{marginBottom:12}}>
<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
<span style={{fontSize:20}}>🚚</span>
<span style={{fontSize:18,fontWeight:800,color:C.wht}}>Delivery Order</span>
</div>
{/* Bar kode Pertamina — memanjang tipis di bawah judul */}
<div style={{background:C.card,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",display:"flex",alignItems:"center",gap:0,flexWrap:"wrap",overflow:"hidden"}}>
{/* Bulan SA */}
{data.company?.saBulan&&<div style={{display:"flex",alignItems:"center",gap:6,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14}}>
<span style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:C.gl2}}>SA Aktif</span>
<span style={{fontSize:11,fontWeight:700,color:C.olt}}>{data.company.saBulan}</span>
</div>}
{/* Kode penebusan */}
{[["Sold To",data.company?.soldTo],["Ship To KCR",data.company?.shipToKCR],["Ship To MGL",data.company?.shipToMGL]].map(([l,v])=>v?<div key={l} style={{display:"flex",alignItems:"center",gap:5,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14}}>
<span style={{fontSize:8,color:C.gl2,fontWeight:600,whiteSpace:"nowrap"}}>{l}:</span>
<span style={{fontSize:11,fontWeight:800,color:C.blt,fontFamily:"'Courier New',monospace"}}>{v}</span>
</div>:null)}
{/* SA per SPPBE */}
{[["KCR",[["12kg",data.company?.sa12KCR],["5,5kg",data.company?.sa55KCR]]],["MGL",[["12kg",data.company?.sa12MGL],["5,5kg",data.company?.sa55MGL]]]].map(([loc,items])=><div key={loc} style={{display:"flex",alignItems:"center",gap:6,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14,flexWrap:"wrap"}}>
<span style={{fontSize:8,fontWeight:700,color:C.olt,letterSpacing:.5}}>{loc}:</span>
{items.map(([l,v])=><div key={l} style={{display:"flex",alignItems:"center",gap:3}}>
<span style={{fontSize:8,color:C.gl2}}>{l}</span>
{v?<span style={{fontSize:11,fontWeight:800,color:C.glt,fontFamily:"'Courier New',monospace"}}>{v}</span>:<span style={{fontSize:9,color:C.gry,fontStyle:"italic"}}>-</span>}
</div>)}
</div>)}
</div>
</div>
<Card>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={onTanggal}/>
<div style={{marginBottom:10}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Trip</label>
<select value={f.trip} onChange={e=>{setF(p=>({...p,trip:e.target.value}));}} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>
{["Trip 1","Trip 2","Trip 3","Trip 4","Trip 5"].map(t=>{var used=(data.doList||[]).filter(d=>d.tanggal===f.tanggal&&d.trip===t);return <option key={t} value={t}>{t}{used.length>0?" (sudah dipakai)":""}</option>;})}
</select>
</div>
<Sel label="SPPBE" value={f.sppbe} onChange={v=>setF(p=>({...p,sppbe:v}))} opts={SPPBE_OPTS}/>
<Sel label="Ukuran" value={f.ukuran} onChange={onUkuran} opts={SIZES}/>
<Inp label="Qty Tabung" type="number" value={f.qty} onChange={v=>setF(p=>({...p,qty:v}))}/>
<Inp label={"HPP/Unit (modal: "+fR(getModal(data,f.ukuran,"Isi",f.tanggal))+")"} type="number" value={f.hppUnit} onChange={v=>setF(p=>({...p,hppUnit:v}))}/>
<Inp label="Total HPP (auto)" value={fR(totalHPP)} ro={true}/>
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/>
<Sel label="Nama Driver" value={f.dibuatOleh} onChange={v=>setF(p=>({...p,dibuatOleh:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).filter(e=>e.aktif).map(e=>({v:e.id,l:e.nama+" ("+e.posisi+")"}))]}/>
</div>
<div style={{padding:"8px 12px",background:C.mode==="dark"?"#0A1A2A":"#DBEAFE",borderRadius:6,fontSize:11,color:C.blt,marginBottom:10}}>ℹ️ HPP/Unit otomatis terisi dari modal terakhir. Anda bisa edit. Setelah simpan, modal akan terupdate menjadi referensi harga untuk penjualan.</div>
<Btn color="green" onClick={save} dis={!f.qty||!f.trip}>💾 Simpan DO</Btn>
</Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Riwayat DO</div>
<RTbl headers={["Tgl","Trip","SPPBE","Ukuran","Qty","HPP/Unit","Total HPP","Driver","Status","Aksi"]} rows={(data.doList||[]).slice(0,100).map(d=>{
var st=d.status||"diterima";
var stBadge=st==="gantung"?<Bdg color="orange">⏳ Gantung</Bdg>:st==="sangkut"?<Bdg color="red">⚠️ Sangkut</Bdg>:<Bdg color="green">✅ Diterima</Bdg>;
var aksiBtn=st==="gantung"?<div style={{display:"flex",gap:3}}>
<button onClick={()=>terimaDO(d)} style={{background:"#15803D",border:"none",borderRadius:5,padding:"3px 7px",color:"white",fontSize:10,fontWeight:700,cursor:"pointer"}}>✅ Terima</button>
<button onClick={()=>sangkutDO(d)} style={{background:"#B45309",border:"none",borderRadius:5,padding:"3px 7px",color:"white",fontSize:10,fontWeight:700,cursor:"pointer"}}>⚠️ Sangkut</button>
<ActBtns onDel={()=>setDelId(d)}/>
</div>:st==="sangkut"?<div style={{display:"flex",gap:3}}>
<button onClick={()=>releaseDO(d)} style={{background:"#1D4ED8",border:"none",borderRadius:5,padding:"3px 7px",color:"white",fontSize:10,fontWeight:700,cursor:"pointer"}}>🔓 Release</button>
<ActBtns onDel={()=>setDelId(d)}/>
</div>:<ActBtns onDel={()=>setDelId(d)}/>;
return[fDs(d.tanggal),<b style={{color:C.wht}}>{d.trip||d.noDO||"-"}</b>,d.sppbe,<Bdg color={d.ukuran==="12 kg"?"blue":d.ukuran==="5.5 kg"?"green":"orange"}>{d.ukuran}</Bdg>,<b style={{color:C.glt}}>{d.qty}</b>,fR(d.hppUnit||0),<b style={{color:C.olt}}>{fR(d.totalHPP||0)}</b>,d.dibuatOleh||"-",stBadge,aksiBtn];})}/>
</Card>
{delId&&<ConfirmDel msg="Hapus DO?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,doList:(d.doList||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// ─── PIUTANG v4 (filter sama dgn Penjualan, cetak invoice BON) ────────────────
function PiutangMod({data,setData,setInv,toast}){
var C=useTheme();
var[openId,setOpenId]=useState(null);var[delId,setDelId]=useState(null);
var[bF,setBF]=useState({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",status:""});
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));
function makeBonInvObj(b){var plg=(data.pelanggan||[]).find(x=>x.id===b.konsumenId);return{noInv:b.noInv||"#HTS/INV/-/-",tanggal:b.tanggal,konsumen:b.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",items:(b.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price)})),total:b.total,metodeBayar:"BON",isBon:true,catatan:b.ket||""};}
function bayar(b){
if(!bF.nominal)return;var nom=Number(bF.nominal);var newSisa=Math.max(0,b.sisaTagihan-nom);var st=newSisa===0?"lunas":"sebagian";
var payRec={id:uid(),tanggal:toDay(),jumlah:nom,metode:bF.metode,bank:bF.metode==="transfer"?bF.bank:"",salesPenerimaId:bF.salesPenerimaId,salesPenerimaNama:salesList.find(e=>e.id===bF.salesPenerimaId)?.nama||""};
var newBon=(data.bon||[]).map(x=>x.id===b.id?{...x,sisaTagihan:newSisa,status:st,pembayaran:[...(x.pembayaran||[]),payRec]}:x);
var newSet=data.setoranSales||[];
if(bF.metode==="cash"&&bF.salesPenerimaId)newSet=[{id:uid(),tanggal:toDay(),salesId:bF.salesPenerimaId,sumber:"piutang",refId:b.id,nominal:nom,disetor:false,konsumen:b.konsumen},...newSet];
setData(d=>({...d,bon:newBon,setoranSales:newSet}));
setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});setOpenId(null);toast("✓ Pembayaran tercatat!");
}
var rows=useMemo(()=>{
return(data.bon||[]).filter(b=>{
if(barFilter.from&&b.tanggal<barFilter.from)return false;
if(barFilter.to&&b.tanggal>barFilter.to)return false;
if(barFilter.salesId&&b.salesId!==barFilter.salesId&&!(b.salesNama||b.sales||'').toLowerCase().includes((salesList.find(e=>e.id===barFilter.salesId)?.nama||'').toLowerCase()))return false;
if(barFilter.konsumen&&!b.konsumen.toLowerCase().includes(barFilter.konsumen.toLowerCase()))return false;
if(barFilter.status&&b.status!==barFilter.status)return false;
return true;
}).map(b=>{var emp=(data.employees||[]).find(e=>e.id===b.salesId);return{...b,salesNama:emp?.nama||b.salesNama||b.sales||"-",dl:dLeft(b.deadline)};});
},[data.bon,data.employees,barFilter]);
var cols=[
{key:"tanggal",label:"Tgl",render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},
{key:"konsumen",label:"Konsumen",render:r=><b style={{color:C.wht,display:"block",minWidth:120}}>{r.konsumen}</b>,filterable:true,width:127},
{key:"salesNama",label:"Sales",filterable:true},
{key:"deadline",label:"Jatuh Tempo",render:r=>r.deadline?<span style={{color:r.dl<0&&r.status!=="lunas"?C.rlt:r.dl<=3&&r.status!=="lunas"?C.olt:C.gl2}}>{fDs(r.deadline)}{r.status!=="lunas"&&r.dl!=null?" ("+(r.dl<0?Math.abs(r.dl)+"h LEWAT":r.dl+"h)"):""}</span>:"-",filterable:false},
{key:"total",label:"Total",render:r=>fR(r.total),filterable:false},
{key:"sisaTagihan",label:"Sisa",render:r=><b style={{color:r.status==="lunas"?C.glt:C.rlt}}>{fR(r.sisaTagihan)}</b>,filterable:false},
{key:"status",label:"Status",render:r=>r.status==="lunas"?<Bdg color="green">Lunas</Bdg>:r.status==="sebagian"?<Bdg color="orange">Sebagian</Bdg>:<Bdg color="red">Belum</Bdg>,filterable:true,filterType:"select",options:[{v:"lunas",l:"Lunas"},{v:"sebagian",l:"Sebagian"},{v:"belum",l:"Belum"}]},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,render:r=><div style={{display:"flex",gap:4}}><button onClick={()=>setInv(makeBonInvObj(r))} title="Cetak Invoice BON" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button>{r.status!=="lunas"&&<button onClick={()=>{setOpenId(r.id);setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:r.salesId||""});}} style={{background:C.grn,border:"none",borderRadius:6,padding:"4px 7px",color:"white",cursor:"pointer",fontSize:12}}>💳</button>}<button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button></div>},
];
var bonActive=(data.bon||[]).filter(b=>b.status!=="lunas");
var totPiutang=bonActive.reduce((a,b)=>a+b.sisaTagihan,0);
return <div>
<STitle icon="💳" children="Piutang / BON"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
<SC label="Total Bon Aktif" value={bonActive.length} icon="📃" color={C.olt}/>
<SC label="Total Piutang" value={fR(totPiutang)} icon="💰" color={C.rlt}/>
<SC label="Bon Lunas" value={(data.bon||[]).filter(b=>b.status==="lunas").length} icon="✅" color={C.glt}/>
</div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Daftar Bon</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Sales" value={barFilter.salesId} onChange={v=>setBarFilter(p=>({...p,salesId:v}))} opts={[{v:"",l:"Semua"},...salesList.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
<Inp label="Konsumen" value={barFilter.konsumen} onChange={v=>setBarFilter(p=>({...p,konsumen:v}))} placeholder="Cari..." style={{marginBottom:0}}/>
<Sel label="Status" value={barFilter.status} onChange={v=>setBarFilter(p=>({...p,status:v}))} opts={[{v:"",l:"Semua"},{v:"belum",l:"Belum"},{v:"sebagian",l:"Sebagian"},{v:"lunas",l:"Lunas"}]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.salesId||barFilter.konsumen||barFilter.status)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",salesId:"",konsumen:"",status:""})}>✕ Reset Filter</Btn></div>}
</div>
<FilterTbl columns={cols} data={rows} empty="Belum ada bon" maxRows={150}/>
</Card>
{openId&&(()=>{var b=(data.bon||[]).find(x=>x.id===openId);if(!b)return null;var paid=(b.pembayaran||[]).reduce((a,p)=>a+p.jumlah,0);return <Card style={{border:"1px solid "+C.blt}}>
<div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}>
<div><div style={{fontWeight:800,color:C.wht,fontSize:15}}>💳 Bayar: {b.konsumen}</div><div style={{fontSize:11,color:C.gl2}}>{fDs(b.tanggal)} · {b.noInv}</div></div>
<Btn sm color="gray" onClick={()=>setOpenId(null)}>Batal</Btn>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}>{[["Total",b.total,C.wht],["Dibayar",paid,C.glt],["Sisa",b.sisaTagihan,C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"6px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
<div><Inp label="Nominal" type="number" value={bF.nominal} onChange={v=>setBF(p=>({...p,nominal:v}))} placeholder={String(b.sisaTagihan)} style={{marginBottom:4}}/><button onClick={()=>setBF(p=>({...p,nominal:String(b.sisaTagihan)}))} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:5,padding:"3px 8px",color:C.blt,cursor:"pointer",fontSize:11}}>Isi Semua</button></div>
<Sel label="Sales Penerima" value={bF.salesPenerimaId} onChange={v=>setBF(p=>({...p,salesPenerimaId:v}))} opts={[{v:"",l:"Admin/Kantor"},...salesList.map(e=>({v:e.id,l:e.nama}))]}/>
</div>
<div style={{display:"flex",gap:6,marginBottom:8}}>{["cash","transfer"].map(m=><button key={m} onClick={()=>setBF(p=>({...p,metode:m}))} style={{background:bF.metode===m?C.blu:C.nav,color:bF.metode===m?"white":C.wht,border:"1px solid "+(bF.metode===m?C.blt:C.bdr),borderRadius:6,padding:"7px 12px",fontWeight:700,fontSize:12,cursor:"pointer",flex:1}}>{m==="cash"?"💵 Cash":"🏦 Transfer"}</button>)}</div>
<Btn color="green" onClick={()=>bayar(b)} dis={!bF.nominal}>💾 Catat Pembayaran</Btn>
</Card>;})()}
{delId&&<ConfirmDel msg="Hapus bon?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,bon:(d.bon||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// ─── PAYROLL v4 (NO O/X toggle in editor, role-aware default rows) ────────────
function PayrollMod({data,setData,toast}){
var C=useTheme();
var[bulan,setBulan]=useState(toMonth());
var[viewSlip,setViewSlip]=useState(null);
var[editSlip,setEditSlip]=useState(null);
var emps=sortEmp((data.employees||[]).filter(e=>e.aktif));

function buildSlipRows(emp){
var r=calcPayrollFull(emp,bulan,data);
var rows=[];
// ── Penghasilan — semua baris selalu ada (sesuai referensi) ──
rows.push({id:uid(),section:"penghasilan",label:"Gaji Pokok",detail:"",jumlah:r.gajiPokok});
// Intensif — selalu tampil untuk sales (otomatis rekap tabung bulan ini)
rows.push({id:uid(),section:"penghasilan",label:"Intensif 12kg",detail:r.bonus.q12+" Tabung × "+fR(r.bonus.r12),jumlah:r.bonus.b12});
rows.push({id:uid(),section:"penghasilan",label:"Intensif 5,5kg",detail:r.bonus.q55+" Tabung × "+fR(r.bonus.r55),jumlah:r.bonus.b55});
// Bonus — selalu ada, default 0 (editable manager/direktur)
rows.push({id:uid(),section:"penghasilan",label:"Bonus",detail:"",jumlah:0});
// Uang Makan — rekap dari operasional
rows.push({id:uid(),section:"penghasilan",label:"Uang Makan",detail:r.hariHadir+" hari × "+fR(emp.uangMakan||15000),jumlah:r.uangMakan});
// Uang Bongkar DO — khusus Checker (tampil selalu agar manager bisa isi)
rows.push({id:uid(),section:"penghasilan",label:"Uang Bongkar DO",detail:r.bongkarCount+" × Rp 50.000",jumlah:r.bongkarTotal});
// Uang Do SPBBE — khusus Driver Truck
rows.push({id:uid(),section:"penghasilan",label:"Uang Do SPBBE",detail:r.spbbeCount+" trip",jumlah:r.spbbeTotal});
// ── Potongan ──
rows.push({id:uid(),section:"potongan",label:"Potongan Absensi",detail:r.absen+" hari × "+fR(r.absen>0?Math.round(r.gajiPokok/r.totalHariKerja):0),jumlah:0});
rows.push({id:uid(),section:"potongan",label:"Total Pinjaman",detail:fR(r.pinjamanSaldo),jumlah:0,kind:"info"});
rows.push({id:uid(),section:"potongan",label:"Potongan Pinjaman",detail:"",jumlah:0});
// ── Yang sudah diterima ──
rows.push({id:uid(),section:"ydt",label:"Uang Makan",detail:"diambil tengah bulan",jumlah:0});
return{empId:emp.id,nama:emp.nama,posisi:emp.posisi,alamat:emp.alamat||"",telepon:emp.telepon||"",bulan,tanggal:toDay(),hariHadir:r.hariHadir,totalHariKerja:r.totalHariKerja,absen:r.absen,totalPinjaman:r.pinjamanSaldo,potonganPinjaman:0,rows};
}

function SlipEditor({slip,onClose,onSave}){
var[s,setS]=useState(JSON.parse(JSON.stringify(slip)));
function rR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
var totPgh=s.rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=s.rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=s.rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var sisa=Number(s.totalPinjaman||0)-Number(s.potonganPinjaman||0);
var totDiterima=totPgh-totPot-totYdt;
function updRow(id,k,v){setS(p=>({...p,rows:p.rows.map(r=>r.id===id?{...r,[k]:v}:r)}));}
function addRow(section){setS(p=>({...p,rows:[...p.rows,{id:uid(),section,label:"Item baru",detail:"",jumlah:0}]}));}
function delRow(id){setS(p=>({...p,rows:p.rows.filter(r=>r.id!==id)}));}
function updPotPinjaman(v){var n=Number(v)||0;setS(p=>({...p,potonganPinjaman:n,rows:p.rows.map(r=>r.label==="Potongan Pinjaman"?{...r,jumlah:n}:r)}));}
return <Modal title={"✏️ Edit Kwitansi — "+s.nama} onClose={onClose} width={800} saveLabel="💾 Simpan & Cetak" onSave={()=>onSave({...s})}>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
<Inp label="Total Hari Kerja" type="number" value={s.totalHariKerja} onChange={v=>setS(p=>({...p,totalHariKerja:Number(v)}))}/>
<Inp label="Hadir" type="number" value={s.hariHadir} onChange={v=>setS(p=>({...p,hariHadir:Number(v)}))}/>
<Inp label="Absen" type="number" value={s.absen} onChange={v=>setS(p=>({...p,absen:Number(v)}))}/>
</div>
{[["penghasilan","Penghasilan",C.glt],["potongan","Potongan",C.rlt],["ydt","Yang sudah diterima",C.olt]].map(x=><div key={x[0]} style={{marginBottom:14,border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden"}}>
<div style={{padding:"8px 12px",background:C.nav,display:"flex",justifyContent:"space-between",alignItems:"center"}}><b style={{color:x[2],fontSize:13}}>{x[1]}</b><Btn sm color="blue" onClick={()=>addRow(x[0])}>+ Baris</Btn></div>
{s.rows.filter(r=>r.section===x[0]).length===0&&<div style={{padding:12,color:C.gl2,fontSize:12,textAlign:"center"}}>Belum ada baris</div>}
{s.rows.filter(r=>r.section===x[0]).map(r=><div key={r.id} style={{display:"grid",gridTemplateColumns:"1fr 150px 130px 24px",gap:6,padding:"6px 10px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}>
<input value={r.label} onChange={e=>updRow(r.id,"label",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 8px",color:C.wht,fontSize:12,outline:"none"}}/>
<input value={r.detail} onChange={e=>updRow(r.id,"detail",e.target.value)} placeholder="@" style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 8px",color:C.wht,fontSize:12,outline:"none"}}/>
<input type="number" value={r.jumlah} onChange={e=>updRow(r.id,"jumlah",Number(e.target.value))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 8px",color:C.wht,fontSize:12,outline:"none"}}/>
<button onClick={()=>delRow(r.id)} title="Hapus" style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13}}>✕</button>
</div>)}
</div>)}
<Card style={{marginBottom:10,border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:8,fontSize:13}}>💰 Sistem Pinjaman Berjalan</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
<Inp label="Total Pinjaman (saldo)" type="number" value={s.totalPinjaman} onChange={v=>setS(p=>({...p,totalPinjaman:Number(v)}))}/>
<Inp label="Potongan Bulan Ini" type="number" value={s.potonganPinjaman} onChange={updPotPinjaman}/>
<div><label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Sisa Pinjaman (auto)</label><div style={{background:"#FFF35B",color:"#1a1a1a",padding:"9px 11px",borderRadius:8,fontWeight:800,fontSize:14}}>{rR(sisa)}</div></div>
</div>
</Card>
<div style={{background:C.nav,borderRadius:10,padding:"14px 16px",border:"1px solid "+C.glt}}>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><span style={{color:C.gl2}}>Total Penghasilan (+)</span><b style={{color:C.glt}}>{rR(totPgh)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><span style={{color:C.gl2}}>Total Potongan (−)</span><b style={{color:C.rlt}}>{rR(totPot)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}><span style={{color:C.gl2}}>Yang sudah diterima (−)</span><b style={{color:C.olt}}>{rR(totYdt)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid "+C.glt,paddingTop:8}}><b style={{color:C.wht,fontSize:14}}>TOTAL DITERIMA</b><b style={{color:C.glt,fontSize:18}}>{rR(totDiterima)}</b></div>
</div>
</Modal>;
}

return <div>
<STitle icon="💼" children="Payroll & Kwitansi Slip Gaji"/>
<Card><MonthPicker label="Pilih Bulan Payroll" value={bulan} onChange={setBulan}/></Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Karyawan — {bulan}</div>
<RTbl headers={["Nama","Posisi","Hadir","Pinjaman","Aksi"]} rows={emps.map(e=>{var r=calcPayrollFull(e,bulan,data);return[<b style={{color:C.wht}}>{e.nama}</b>,e.posisi,<Bdg color="blue">{r.hariHadir}/{r.totalHariKerja}</Bdg>,<b style={{color:r.pinjamanSaldo>0?C.olt:C.gl2}}>{fR(r.pinjamanSaldo)}</b>,<Btn sm color="blue" onClick={()=>setEditSlip(buildSlipRows(e))}>📝 Buat Slip</Btn>];})}/>
</Card>
{(data.payrollLog||[]).length>0&&<Card><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📜 Riwayat Slip</div><div style={{display:"flex",gap:6}}><Bdg color="green">Aktif: {(data.payrollLog||[]).filter(p=>!p.arsip).length}</Bdg><Bdg color="gray">Arsip: {(data.payrollLog||[]).filter(p=>p.arsip).length}</Bdg></div></div><RTbl headers={["Bulan","Nama","No.Slip","Total","Status","Aksi"]} rows={(data.payrollLog||[]).map(p=>[
p.bulan,p.nama,p.noSlip,
<b style={{color:C.glt}}>{fR(p.totalDiterima)}</b>,
p.arsip?<Bdg color="gray">Diarsip</Bdg>:<Bdg color="green">Aktif</Bdg>,
<div style={{display:"flex",gap:4}}>
<Btn sm color="blue" onClick={()=>setViewSlip(p)}>🖨️ Cetak</Btn>
{!p.arsip
?<Btn sm color="gray" onClick={()=>setData(d=>({...d,payrollLog:(d.payrollLog||[]).map(x=>x.id===p.id?{...x,arsip:true}:x)}))}>📦 Arsip</Btn>
:<Btn sm color="green" onClick={()=>setData(d=>({...d,payrollLog:(d.payrollLog||[]).map(x=>x.id===p.id?{...x,arsip:false}:x)}))}>↩️ Restore</Btn>}
</div>
])}/></Card>}
{editSlip&&<SlipEditor slip={editSlip} onClose={()=>setEditSlip(null)} onSave={s=>{
var sgInfo=nextSGNo(data,s.bulan);
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.sg)newCounters.sg={};newCounters.sg[sgInfo.key]=sgInfo.n;
var totPgh=s.rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=s.rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=s.rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totalDiterima=totPgh-totPot-totYdt;
var rec={id:uid(),noSlip:sgInfo.no,empId:s.empId,nama:s.nama,posisi:s.posisi,alamat:s.alamat,telepon:s.telepon,bulan:s.bulan,tanggal:s.tanggal,hariHadir:s.hariHadir,totalHariKerja:s.totalHariKerja,absen:s.absen,totalPinjaman:s.totalPinjaman,potonganPinjaman:s.potonganPinjaman,rows:s.rows,totalDiterima};
setData(d=>({...d,payrollLog:[rec,...(d.payrollLog||[])],counters:newCounters}));
setEditSlip(null);setViewSlip(rec);toast("✓ Slip "+sgInfo.no+" tersimpan!");
}}/>}
{viewSlip&&<SlipGajiView slip={viewSlip} company={data.company} onClose={()=>setViewSlip(null)}/>}
</div>;
}

// ─── PELANGGAN ────────────────────────────────────────────────────────────────
function PelangganMod({data,setData,toast}){
var C=useTheme();var mob=useMobile();
var blk={nama:"",kategori:"Rumah Tangga",kategoriCustom:"",telepon:"",alamat:""};
var[f,setF]=useState({...blk});var[edit,setEdit]=useState(null);var[delId,setDelId]=useState(null);
var[search,setSearch]=useState("");var[filterKat,setFilterKat]=useState("");
var[hargaModal,setHargaModal]=useState(null);var[hf,setHf]=useState({ukuran:"12 kg",jenis:"Isi",harga:""});
var[regEdit,setRegEdit]=useState(null);
var isCustom=f.kategori==="Lainnya";
var allKats=[...new Set([...PLG_KAT_27,"Lainnya",...(data.pelanggan||[]).map(p=>p.kategori).filter(k=>!PLG_KAT_27.includes(k))])];
var filtered=(data.pelanggan||[]).filter(p=>{if(filterKat&&p.kategori!==filterKat)return false;if(search&&!((p.nama||"")+" "+(p.telepon||"")+" "+(p.regNo||"")).toLowerCase().includes(search.toLowerCase()))return false;return true;});
function getKat(){return isCustom?(f.kategoriCustom||"Lainnya"):f.kategori;}
function save(){if(!f.nama)return;if(edit){setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id===edit.id?{...p,...f,kategori:getKat()}:p)}));setEdit(null);}else{var regInfo=nextRegNo(data);var newCounters={...(data.counters||{inv:{},sg:{},reg:0}),reg:regInfo.n};setData(d=>({...d,pelanggan:[{id:uid(),...f,kategori:getKat(),regNo:regInfo.no,hargaKhusus:[]},...(d.pelanggan||[])],counters:newCounters}));}setF({...blk});toast("✓ Pelanggan disimpan!");}
function saveHarga(plg){if(!hf.harga)return;var baseHK=Array.isArray(plg.hargaKhusus)?plg.hargaKhusus:[];var newHK=[...baseHK.filter(x=>!(x.ukuran===hf.ukuran&&x.jenis===hf.jenis)),{ukuran:hf.ukuran,jenis:hf.jenis,harga:Number(hf.harga)}];setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id!==plg.id?p:{...p,hargaKhusus:newHK})}));setHargaModal(prev=>prev?{...prev,hargaKhusus:newHK}:null);toast("✓ Harga khusus disimpan!");}
function delHarga(plg,ukuran,jenis){var newHK=(plg.hargaKhusus||[]).filter(x=>!(x.ukuran===ukuran&&x.jenis===jenis));setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id!==plg.id?p:{...p,hargaKhusus:newHK})}));setHargaModal(prev=>prev?{...prev,hargaKhusus:newHK}:null);}
return <div>
<STitle icon="👥" children="Kelola Pelanggan"/>
<Card>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10}}>
<Inp label="Nama Pelanggan" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))} style={{gridColumn:mob?"1/-1":"auto"}}/>
<div style={{marginBottom:10,gridColumn:mob?"1/-1":"auto"}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label>
<select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{[...PLG_KAT_27,"Lainnya"].map(k=><option key={k}>{k}</option>)}</select>
{isCustom&&<input value={f.kategoriCustom} placeholder="Ketik kategori..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}
</div>
<Inp label="No. Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/>
<Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/>
</div>
<Btn color="green" onClick={save} dis={!f.nama}>➕ {edit?"Update":"Tambah"} Pelanggan</Btn>
{!edit&&<div style={{fontSize:11,color:C.gl2,marginTop:6}}>No. Registrasi otomatis: <b style={{color:C.blt}}>HTS/CST/{String((data.counters?.reg||0)+1).padStart(3,"0")}</b></div>}
</Card>
<Card>
<div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,marginBottom:10}}><Inp label="" value={search} onChange={setSearch} placeholder="🔍 Cari nama / HP / No.Reg..." style={{marginBottom:0}}/><Sel label="" value={filterKat} onChange={setFilterKat} opts={[{v:"",l:"Semua"},...allKats.map(k=>({v:k,l:k}))]} style={{marginBottom:0}}/></div>
<div style={{fontSize:12,color:C.gl2,marginBottom:8}}>{filtered.length} pelanggan</div>
<RTbl headers={["No.Reg","Nama","Kategori","Telepon","Harga Khusus","Aksi"]} rows={filtered.map(p=>{var hk=p.hargaKhusus||[];return[<button onClick={()=>setRegEdit(p)} style={{background:"none",border:"none",color:C.blt,fontSize:11,fontWeight:700,cursor:"pointer",padding:0}}>{p.regNo||"-"}</button>,<b style={{color:C.wht}}>{p.nama}</b>,<Bdg color={PLG_TITIP_KAT.includes(p.kategori)?"blue":"gray"}>{p.kategori||"-"}</Bdg>,p.telepon||"-",<button onClick={()=>{var cur=(data.pelanggan||[]).find(x=>x.id===p.id)||p;setHargaModal({...cur,hargaKhusus:Array.isArray(cur.hargaKhusus)?cur.hargaKhusus:[]});}} style={{background:hk.length>0?C.inHv:C.nav,border:"1px solid "+(hk.length>0?C.blt:C.bdr),borderRadius:5,padding:"3px 7px",color:hk.length>0?C.blt:C.gl2,cursor:"pointer",fontSize:11}}>{hk.length>0?hk.length+" harga":"+ Tambah"}</button>,<ActBtns onEdit={()=>{setEdit(p);setF({nama:p.nama,kategori:PLG_KAT_27.includes(p.kategori)?p.kategori:"Lainnya",kategoriCustom:PLG_KAT_27.includes(p.kategori)?"":p.kategori,telepon:p.telepon||"",alamat:p.alamat||""});}} onDel={()=>setDelId(p)}/>];})}/>
</Card>
{regEdit&&<Modal title={"Edit No. Registrasi — "+regEdit.nama} onClose={()=>setRegEdit(null)} onSave={()=>{setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id===regEdit.id?{...p,regNo:regEdit.regNo}:p)}));setRegEdit(null);toast("✓ Diperbarui!");}}><Inp label="No. Registrasi" value={regEdit.regNo} onChange={v=>setRegEdit(p=>({...p,regNo:v}))} placeholder="HTS/CST/001"/></Modal>}
{hargaModal&&<Modal title={"💲 Harga Khusus — "+hargaModal.nama} onClose={()=>setHargaModal(null)} width={500}>
<div style={{marginBottom:14}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>Harga Terdaftar:</div>
{!(hargaModal.hargaKhusus||[]).length&&<div style={{color:C.gl2,fontSize:12,marginBottom:10,fontStyle:"italic"}}>Belum ada. Pakai HET default.</div>}
{(hargaModal.hargaKhusus||[]).map((h,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:C.nav,borderRadius:8,marginBottom:6}}>
<span style={{fontSize:13,color:C.wht}}><b>{h.ukuran}</b> — {h.jenis}</span>
<div style={{display:"flex",alignItems:"center",gap:8}}><span style={{color:C.glt,fontWeight:700}}>{fR(h.harga)}</span><span style={{color:C.gl2,fontSize:10}}>(HET: {fR(getHET(data,h.ukuran,h.jenis))})</span><button onClick={()=>delHarga(hargaModal,h.ukuran,h.jenis)} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:11,padding:"3px 7px"}}>🗑️</button></div>
</div>)}
</div>
<div style={{borderTop:"1px solid "+C.bdr,paddingTop:14}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>Tambah / Update:</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
<Sel label="Ukuran" value={hf.ukuran} onChange={v=>setHf(p=>({...p,ukuran:v}))} opts={SIZES}/>
<Sel label="Jenis" value={hf.jenis} onChange={v=>setHf(p=>({...p,jenis:v}))} opts={JENIS}/>
<Inp label={"Harga (HET: "+fR(getHET(data,hf.ukuran,hf.jenis))+")"} type="number" value={hf.harga} onChange={v=>setHf(p=>({...p,harga:v}))}/>
</div>
<Btn color="green" onClick={()=>{saveHarga(hargaModal);setHf({ukuran:"12 kg",jenis:"Isi",harga:""});}} dis={!hf.harga}>💾 Simpan Harga</Btn>
</div>
</Modal>}
{edit&&<Modal title={"Edit: "+edit.nama} onSave={save} onClose={()=>{setEdit(null);setF({...blk});}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><div style={{marginBottom:10}}><label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label><select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{[...PLG_KAT_27,"Lainnya"].map(k=><option key={k}>{k}</option>)}</select>{isCustom&&<input value={f.kategoriCustom} placeholder="Ketik..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}</div><Inp label="Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/></Modal>}
{delId&&<ConfirmDel msg={"Hapus \""+delId.nama+"\"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,pelanggan:(d.pelanggan||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// === AKHIR BAGIAN 3 ===
// === BAGIAN 4 DARI 4 ===

// ─── PENGELUARAN v4 (fixed + filter) ──────────────────────────────────────────
function PengeluaranMod({data,setData,user,toast}){
var C=useTheme();
var[f,setF]=useState({tanggal:toDay(),kategori:"BBM",kategoriCustom:"",keperluan:"perusahaan",ket:"",nominal:"",metode:"cash",bank:"BSI"});
var[delId,setDelId]=useState(null);
var[barFilter,setBarFilter]=useState({from:"",to:"",kategori:"",keperluan:""});
var karList=sortEmp((data.employees||[]).filter(e=>e.aktif));
var isLainnya=f.kategori==="Lainnya";
function save(){
if(!f.nominal||Number(f.nominal)<=0)return;
var kat=isLainnya?(f.kategoriCustom||"Lainnya"):f.kategori;
var isKar=f.keperluan!=="perusahaan";
var karId=isKar?f.keperluan:null;
var karNm=isKar?(karList.find(e=>e.id===karId)?.nama||""):"";
var rec={id:uid(),tanggal:f.tanggal,kategori:kat,keperluan:f.keperluan,karyawanId:karId,karyawanNama:karNm,ket:f.ket,nominal:Number(f.nominal),metode:f.metode,bank:f.metode==="transfer"?f.bank:""};
setData(d=>({...d,pengeluaran:[rec,...(d.pengeluaran||[])]}));
setF(p=>({...p,nominal:"",ket:""}));
toast("✓ Pengeluaran dicatat!");
}
var penFiltered=useMemo(()=>{
return(data.pengeluaran||[]).filter(p=>{
if(barFilter.from&&(p.tanggal||"")<barFilter.from)return false;
if(barFilter.to&&(p.tanggal||"")>barFilter.to)return false;
if(barFilter.kategori&&(p.kategori||"")!==barFilter.kategori)return false;
if(barFilter.keperluan){
if(barFilter.keperluan==="perusahaan"&&p.keperluan!=="perusahaan")return false;
if(barFilter.keperluan!=="perusahaan"&&p.karyawanId!==barFilter.keperluan)return false;
}
return true;
});
},[data.pengeluaran,barFilter]);
var totalPen=penFiltered.reduce((a,p)=>a+Number(p.nominal||0),0);
var cashPen=penFiltered.filter(p=>p.metode==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var tfPen=penFiltered.filter(p=>p.metode==="transfer").reduce((a,p)=>a+Number(p.nominal||0),0);
var cols=[
{key:"tanggal",label:"Tgl",render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},
{key:"kategori",label:"Kategori",render:r=><Bdg color="gray">{r.kategori}</Bdg>,filterable:true},
{key:"keperluan",label:"Keperluan",render:r=>r.keperluan==="perusahaan"?<Bdg color="gray">Perusahaan</Bdg>:<Bdg color="blue">{r.karyawanNama||r.keperluan}</Bdg>,filterable:true},
{key:"ket",label:"Keterangan",render:r=>r.ket||"-",filterable:true},
{key:"nominal",label:"Nominal",render:r=><b style={{color:C.rlt}}>{fR(r.nominal)}</b>,sortVal:r=>r.nominal,filterable:false},
{key:"metode",label:"Metode",render:r=><Bdg color={r.metode==="transfer"?"blue":"green"}>{r.metode}</Bdg>,filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"}]},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,render:r=><button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button>},
];
return <div>
<STitle icon="💸" children="Pengeluaran"/>
<Card>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/>
<div style={{marginBottom:10}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label>
<select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{KAT_K.map(k=><option key={k}>{k}</option>)}</select>
{isLainnya&&<input value={f.kategoriCustom} placeholder="Ketik kategori..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}
</div>
<Sel label="Keperluan" value={f.keperluan} onChange={v=>setF(p=>({...p,keperluan:v}))} opts={[{v:"perusahaan",l:"🏢 Perusahaan"},...karList.map(e=>({v:e.id,l:e.nama}))]}/>
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))} placeholder="Detail..."/>
<Inp label="Nominal" type="number" value={f.nominal} onChange={v=>setF(p=>({...p,nominal:v}))}/>
</div>
<div style={{display:"flex",gap:8,marginBottom:10}}>{["cash","transfer"].map(m=><button key={m} onClick={()=>setF(p=>({...p,metode:m}))} style={{background:f.metode===m?C.blu:C.nav,color:f.metode===m?"white":C.wht,border:"1px solid "+(f.metode===m?C.blt:C.bdr),borderRadius:8,padding:"6px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{m==="cash"?"💵 Cash":"🏦 Transfer"}</button>)}</div>
{f.metode==="transfer"&&<div style={{display:"flex",gap:8,marginBottom:10}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,bank:b}))} style={{background:f.bank===b?C.blu:C.nav,color:f.bank===b?"white":C.wht,border:"2px solid "+(f.bank===b?C.blt:C.bdr),borderRadius:8,padding:"6px 14px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
<Btn color="red" onClick={save} dis={!f.nominal||Number(f.nominal)<=0}>💾 Catat Pengeluaran</Btn>
</Card>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:12}}>
<SC label="Total Pengeluaran" value={fR(totalPen)} icon="💸" color={C.rlt}/>
<SC label="Cash" value={fR(cashPen)} icon="💵" color={C.gl2}/>
<SC label="Transfer" value={fR(tfPen)} icon="🏦" color={C.blt}/>
</div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Pengeluaran</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Kategori" value={barFilter.kategori} onChange={v=>setBarFilter(p=>({...p,kategori:v}))} opts={[{v:"",l:"Semua"},...KAT_K.map(k=>({v:k,l:k}))]} style={{marginBottom:0}}/>
<Sel label="Keperluan" value={barFilter.keperluan} onChange={v=>setBarFilter(p=>({...p,keperluan:v}))} opts={[{v:"",l:"Semua"},{v:"perusahaan",l:"Perusahaan"},...karList.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.kategori||barFilter.keperluan)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",kategori:"",keperluan:""})}>✕ Reset Filter</Btn></div>}
</div>
<FilterTbl columns={cols} data={penFiltered} empty="Belum ada pengeluaran" maxRows={100}/>
</Card>
{delId&&<ConfirmDel msg={"Hapus pengeluaran "+fR(delId.nominal)+"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,pengeluaran:(d.pengeluaran||[]).filter(x=>x.id!==delId.id)}));setDelId(null);toast("✓ Dihapus");}}/>}
</div>;
}

// ─── SETORAN SALES v4 (buat ulang total) ──────────────────────────────────────
function SetoranMod({data,setData,user,toast}){
var C=useTheme();
var isSales=user&&["sales_driver","sales_freelance"].includes(user.role);
var[salesId,setSalesId]=useState(isSales?user.id:"");
var[tgl,setTgl]=useState(toDay());
var[pecah,setPecah]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[jadikanPinjaman,setJadikanPinjaman]=useState(false);
var[viewTB,setViewTB]=useState(null);// untuk lihat detail tutup buku
// Pengeluaran ops hanya catatan rekap, tidak jadi hutang
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));

// Cash masuk dari penjualan hari itu
var penjHari=(data.penjualan||[]).filter(p=>p.salesId===salesId&&p.tanggal===tgl&&p.bayar==="cash");
var cashPenjualan=penjHari.reduce((a,p)=>a+(p.total||0),0);

// Cash masuk dari pembayaran piutang hari itu
var piutangHari=(data.setoranSales||[]).filter(s=>s.salesId===salesId&&s.tanggal===tgl&&s.sumber==="piutang"&&!s.disetor);
var cashPiutang=piutangHari.reduce((a,s)=>a+(s.nominal||0),0);
var totalWajib=cashPenjualan+cashPiutang;

// Pengeluaran ops hari itu — hanya info rekap, tidak dipotong dari setoran
var opsHari=(data.pengeluaran||[]).filter(p=>p.karyawanId===salesId&&p.tanggal===tgl&&KAT_OPS.includes(p.kategori));
var totalOpsHari=opsHari.reduce((a,op)=>a+op.nominal,0);
var bersihSetor=totalWajib;// Setoran tidak dipotong ops
var totalTunai=DENOMS.reduce((a,d)=>a+Number(pecah[d]||0)*d,0);
var selisih=totalTunai-bersihSetor;



function konfirmasi(){
var newSet=(data.setoranSales||[]).map(s=>{if(s.salesId===salesId&&s.tanggal===tgl&&!s.disetor)return{...s,disetor:true,tanggalSetor:tgl};return s;});
var emp=(data.employees||[]).find(e=>e.id===salesId);
var newAmb=[...(data.ambilan||[])];
// Kalau kurang setor dan user pilih jadikan pinjaman
if(selisih<0&&jadikanPinjaman){
newAmb.unshift({id:uid(),karyawanId:salesId,karyawanNama:emp?.nama||"",nominal:Math.abs(selisih),ket:"Kurang setor "+fDs(tgl),tanggal:tgl});
}
var logEntry={id:uid(),tanggal:tgl,salesId,salesNama:emp?.nama||"",cashPenjualan,cashPiutang,totalWajib,bersihSetor,totalTunai,selisih,jadikanPinjaman:selisih<0&&jadikanPinjaman};
setData(d=>({...d,setoranSales:newSet,ambilan:newAmb,setoranLog:[logEntry,...(d.setoranLog||[])]}));
setPecah(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
setJadikanPinjaman(false);
toast("✓ Setoran dikonfirmasi! Selisih: "+fR(selisih));
}

var riwayat=(data.setoranLog||[]).filter(s=>!salesId||s.salesId===salesId).slice(0,20);
return <div>
<STitle icon="💰" children="Setoran Sales"/>
<Card>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
<Sel label="Sales" value={salesId} onChange={v=>setSalesId(v)} opts={[{v:"",l:"-- Pilih --"},...salesList.map(e=>({v:e.id,l:e.nama}))]}/>
<Inp label="Tanggal" type="date" value={tgl} onChange={v=>setTgl(v)}/>
</div>
</Card>
{salesId&&<>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>📊 Ringkasan Wajib Setor — {fDs(tgl)}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}>
{[["Cash Penjualan",cashPenjualan,C.glt,penjHari.length+" trx"],["Cash Piutang",cashPiutang,C.blt,piutangHari.length+" bon"],["TOTAL WAJIB",totalWajib,C.wht,""]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2,marginBottom:2}}>{x[0]}</div><div style={{fontSize:14,fontWeight:900,color:x[2]}}>{fR(x[1])}</div>{x[3]&&<div style={{fontSize:10,color:C.gl2,marginTop:2}}>{x[3]}</div>}</div>)}
</div>
{penjHari.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Detail penjualan cash:</div>{penjHari.map(p=><div key={p.id} style={{display:"flex",justifyContent:"space-between",padding:"3px 8px",background:C.bg,borderRadius:5,marginBottom:2,fontSize:12}}><span style={{color:C.gl2}}>{p.konsumen}</span><b style={{color:C.glt}}>{fR(p.total)}</b></div>)}</div>}
</Card>
{opsHari.length>0&&<Card style={{border:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📋 Pengeluaran Operasional Hari Ini (Info)</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Catatan rekap pengeluaran operasional — tidak dipotong dari setoran</div>
{opsHari.map(op=><div key={op.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 12px",background:C.nav,borderRadius:8,marginBottom:4,border:"1px solid "+C.bdr}}>
<div><div style={{fontSize:12,fontWeight:600,color:C.wht}}>{op.kategori}</div><div style={{fontSize:11,color:C.gl2}}>{op.ket||""}</div></div>
<b style={{color:C.olt}}>{fR(op.nominal)}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",padding:"7px 12px",background:C.bg,borderRadius:6,border:"1px solid "+C.bdr,marginTop:4}}>
<span style={{color:C.gl2,fontSize:12}}>Total ops hari ini</span>
<b style={{color:C.olt}}>{fR(totalOpsHari)}</b>
</div>
</Card>}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💵 Pecahan Kas Fisik yang Disetor</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{DENOMS.map(d=><div key={d} style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"5px 11px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}>
<span style={{color:C.wht,fontSize:13,fontWeight:600}}>{fR(d)}</span>
<input type="number" value={pecah[d]} placeholder="0" onChange={e=>setPecah(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.wht,fontSize:12,outline:"none",width:74}}/>
<span style={{color:Number(pecah[d]||0)>0?C.olt:C.gl2,fontWeight:700,fontSize:12}}>{Number(pecah[d]||0)>0?fR(Number(pecah[d])*d):"-"}</span>
</div>)}
<div style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"9px 11px",background:C.nav,borderTop:"2px solid "+C.bdr}}><b style={{color:C.wht}}>Total Tunai</b><span/><b style={{color:C.glt}}>{fR(totalTunai)}</b></div>
</div>
</Card>
<Card style={{border:"2px solid "+(Math.abs(selisih)<1000?C.glt:C.rlt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔄 Rekonsiliasi</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
{[["Total Wajib Setor",totalWajib,C.wht],["Total Tunai Fisik",totalTunai,C.glt],["Selisih",selisih,selisih>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:Math.abs(selisih)<1000?C.grn:C.rdk,borderRadius:8,marginBottom:12}}>
<span style={{color:"white",fontWeight:700}}>SELISIH {selisih>=0?"LEBIH":"KURANG"}</span>
<b style={{color:"white",fontSize:18}}>{fR(Math.abs(selisih))}</b>
</div>
<div style={{marginBottom:8}}>
{selisih<0&&<div style={{marginBottom:8}}>
<label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",padding:"8px 12px",background:C.rdk,borderRadius:8}}>
<input type="checkbox" checked={jadikanPinjaman} onChange={e=>setJadikanPinjaman(e.target.checked)} style={{width:16,height:16}}/>
<span style={{color:"white",fontSize:12,fontWeight:700}}>💳 Jadikan pinjaman karyawan (kurang setor {fR(Math.abs(selisih))})</span>
</label>
</div>}
</div>
<Btn color="green" onClick={konfirmasi} dis={!salesId}>✓ Konfirmasi Setoran</Btn>
</Card>
</>}
{riwayat.length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Setoran</div>
<RTbl headers={["Tgl","Sales","Wajib Setor","Tunai","Selisih"]} rows={riwayat.map(r=>[fDs(r.tanggal),r.salesNama,fR(r.totalWajib),fR(r.totalTunai),<b style={{color:Math.abs(r.selisih||0)<1000?C.glt:C.rlt}}>{fR(r.selisih||0)}</b>])}/>
</Card>}
</div>;
}

// ─── TUTUP BUKU v4 (harian + bulanan, PDF/PNG) ────────────────────────────────
function TabelStokHarian({data,tgl}){
var C=useTheme();
var prevTB=(data.tutupBuku||[]).filter(r=>r.tanggal&&r.tanggal<tgl).sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0];
var rows=SIZES.map(s=>{
  var stokAwal=prevTB?.detail?.stokSnapshot?.[s]?.isi??((data.stock||{})[s]||0);
  var doMasuk=(data.doList||[]).filter(d=>d.tanggal===tgl&&d.ukuran===s&&(d.status||"diterima")==="diterima").reduce((a,d)=>a+Number(d.qty||0),0);
  var terjual=(data.penjualan||[]).filter(p=>p.tanggal===tgl).reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran===s).reduce((b,it)=>b+Number(it.qty||0),0),0);
  var sisaAkhir=(data.stock||{})[s]||0;
  var totalHariIni=stokAwal+doMasuk;
  return{s,stokAwal,doMasuk,totalHariIni,terjual,sisaAkhir};
});
return <div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"inherit"}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"8px 12px",color:C.gl2,fontWeight:700,textAlign:"left",borderBottom:"2px solid "+C.bdr,fontSize:11}}>Keterangan</th>
{SIZES.map(s=><th key={s} style={{padding:"8px 12px",color:C.wht,fontWeight:700,textAlign:"center",borderBottom:"2px solid "+C.bdr,fontSize:11}}>{s}</th>)}
</tr></thead>
<tbody>
{[
["Stok Awal",rows.map(r=>r.stokAwal),C.gl2,false],
["DO Masuk (Diterima)",rows.map(r=>r.doMasuk),C.blt,false],
["Total Stok Harian",rows.map(r=>r.totalHariIni),C.wht,true],
["Penjualan",rows.map(r=>r.terjual),C.rlt,false],
["SISA STOK AKHIR",rows.map(r=>r.sisaAkhir),C.glt,true],
].map((row,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:row[3]?C.nav:"transparent"}}>
<td style={{padding:"7px 12px",color:row[3]?C.wht:C.gl2,fontWeight:row[3]?700:400}}>{row[0]}</td>
{row[1].map((v,j)=><td key={j} style={{padding:"7px 12px",textAlign:"center",fontWeight:row[3]?800:600,fontSize:row[3]?14:12,color:row[2]}}>{v}</td>)}
</tr>)}
</tbody>
</table>
</div>;
}
function TutupBukuMod({data,setData,toast}){
var C=useTheme();var[tab,setTab]=useState("harian");
var[tgl,setTgl]=useState(toDay());var[bln,setBln]=useState(toMonth());
var[cashLaci,setCashLaci]=useState("");var[rekBSI,setRekBSI]=useState("");var[rekBCA,setRekBCA]=useState("");
var[pecah,setPecah]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[jadikanPinjaman,setJadikanPinjaman]=useState(false);
var[viewTB,setViewTB]=useState(null);// untuk lihat detail tutup buku

// Harian P&L
var pHari=(data.penjualan||[]).filter(e=>e.tanggal===tgl);
var omzetH=pHari.reduce((a,e)=>a+(e.total||0),0);
var marginH=pHari.reduce((a,e)=>a+(e.margin||0),0);
var hppH=omzetH-marginH;
var penH=(data.pengeluaran||[]).filter(e=>e.tanggal===tgl);
var totalOutH=penH.reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersihH=marginH-totalOutH;
var cashInH=pHari.filter(e=>e.bayar==="cash").reduce((a,e)=>a+(e.total||0),0);
var tfInH=pHari.filter(e=>e.bayar==="transfer").reduce((a,e)=>a+(e.total||0),0);
var bonInH=pHari.filter(e=>e.bayar==="bon").reduce((a,e)=>a+(e.total||0),0);

// Asset kalkulasi
var piutangA=calcTotalPiutang(data);
var nilaiStokA=calcNilaiStok(data);
var pinjamanA=Math.max(0,calcPinjamanKaryawan(data));
var kosong55=getKosong(data,"5.5 kg");var kosong12=getKosong(data,"12 kg");var kosong50=getKosong(data,"50 kg");
var totalPecah=DENOMS.reduce((a,d)=>a+Number(pecah[d]||0)*d,0);
var cashTotal=(Number(cashLaci)||0)+(Number(rekBSI)||0)+(Number(rekBCA)||0);

// DO Gantung = DO belum diklik Diterima/Sangkut
var doGantung=(data.doList||[]).filter(d=>d.status==="gantung");
var nilaiDOGantung=doGantung.reduce((a,d)=>a+Number(d.totalHPP||0),0);
var doSangkut=(data.doList||[]).filter(d=>d.status==="sangkut");
var nilaiDOSangkut=doSangkut.reduce((a,d)=>a+Number(d.totalHPP||0),0);

// Asset Tabung Milik PT
var titipLuarBal={};
(data.titipList||[]).forEach(t=>{if(t.tipe==="titip_luar"||t.tipe==="tarik_luar"){var m=t.tipe==="titip_luar"?1:-1;(t.items||[]).forEach(it=>{titipLuarBal[it.ukuran]=(titipLuarBal[it.ukuran]||0)+m*Number(it.qty||0);});}});

var hargaTbg=data.company?.hargaTbgKosong||{};
var assetTabungMilikPT=SIZES.reduce((a,s)=>{
  var totalTabung=(data.totalAsset||{})[s]||0;
  var titipLuar=Math.max(0,titipLuarBal[s]||0);
  var milikPT=Math.max(0,totalTabung-titipLuar);
  return a+milikPT*(hargaTbg[s]||0);
},0);
var assetArmada=Number(data.company?.assetArmada)||0;

// PIUTANG & MODAL BERJALAN = Stok Isi + DO Gantung + BON Konsumen + Pinjaman Karyawan
var piutangModal=nilaiStokA+nilaiDOGantung+piutangA+pinjamanA;

// TOTAL CASH FLOW = Total Cash + Bank + Piutang & Modal Berjalan
var cashFlowOmset=cashTotal+piutangModal;

// Total Asset
var assetValue=cashFlowOmset+assetTabungMilikPT+assetArmada;

// Prev tutup buku
var prevTB=(data.tutupBuku||[]).filter(r=>tab==="harian"?r.tanggal<tgl:r.bulan&&r.bulan<bln).sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0];
var cashFlowKemarin=prevTB?.cashFlowOmset||0;
var deltaSelisih=cashFlowOmset-cashFlowKemarin-(labaBersihH||0);

// Pemasukan lainnya
var[pemasukanLain,setPemasukanLain]=useState("");

// Bulanan aggregation
var pBln=(data.penjualan||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var omzetB=pBln.reduce((a,e)=>a+(e.total||0),0);
var marginB=pBln.reduce((a,e)=>a+(e.margin||0),0);
var penB=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var totalOutB=penB.reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersihB=marginB-totalOutB;
var dim=daysInMonth(bln);
var doBln=(data.doList||[]).filter(e=>(e.tanggal||"").startsWith(bln)&&(e.status||"diterima")==="diterima");
var hppBln=doBln.reduce((a,e)=>a+Number(e.totalHPP||0),0);
var cashFlowKumul=0;
var chartDataB=[];for(var d=1;d<=dim;d++){var ds=bln+"-"+String(d).padStart(2,"0");var pp=pBln.filter(x=>x.tanggal===ds);var oz=pp.reduce((a,x)=>a+(x.total||0),0);var mg=pp.reduce((a,x)=>a+(x.margin||0),0);var pn=penB.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.nominal||0),0);var hpp=doBln.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0);var labaBersih=mg-pn;cashFlowKumul+=labaBersih;chartDataB.push({hari:String(d),omzet:oz,hpp,marginKotor:mg,pengeluaran:pn,labaBersih,cashFlow:cashFlowKumul});}
// Top 5 hari omzet
var top5=chartDataB.slice().sort((a,b)=>b.omzet-a.omzet).slice(0,5);
// Pengeluaran per kategori bulanan
var katPenMap={};penB.forEach(p=>{katPenMap[p.kategori]=(katPenMap[p.kategori]||0)+Number(p.nominal||0);});
var katPenArr=Object.entries(katPenMap).sort((a,b)=>b[1]-a[1]);

function saveHarian(){
var penjualanDetail=(data.penjualan||[]).filter(e=>e.tanggal===tgl).map(e=>({noInv:e.noInv,konsumen:e.konsumen,total:e.total,bayar:e.bayar,items:e.items}));
var pengeluaranDetail=(data.pengeluaran||[]).filter(e=>e.tanggal===tgl).map(e=>({kategori:e.kategori,nominal:e.nominal,ket:e.ket,karyawan:e.karyawanNama||""}));
var doDetail=(data.doList||[]).filter(e=>e.tanggal===tgl).map(e=>({trip:e.trip,sppbe:e.sppbe,ukuran:e.ukuran,qty:e.qty,totalHPP:e.totalHPP,status:e.status}));
var stokSnap={};SIZES.forEach(s=>{stokSnap[s]={isi:(data.stock||{})[s]||0,kosong:getKosong(data,s),titip:getTitipTotal(data.titipList,s)};});
var katPenH={};penH.forEach(p=>{katPenH[p.kategori]=(katPenH[p.kategori]||0)+Number(p.nominal||0);});
var rec={
  id:uid(),tanggal:tgl,
  omzet:omzetH,hpp:hppH,labaKotor:marginH,totalOut:totalOutH,
  pemasukanLain:Number(pemasukanLain)||0,
  labaBersih:labaBersihH+(Number(pemasukanLain)||0),
  cashIn:cashInH,tfIn:tfInH,bonIn:bonInH,
  cashLaci:Number(cashLaci)||0,rekBSI:Number(rekBSI)||0,rekBCA:Number(rekBCA)||0,totalPecah,
  piutangA,nilaiStokA,pinjamanA,
  nilaiDOGantung,nilaiDOSangkut,
  cashFlowOmset,assetTabungMilikPT,assetArmada,assetValue,
  cashFlowKemarin,deltaSelisih,
  detail:{penjualan:penjualanDetail,pengeluaran:pengeluaranDetail,doMasuk:doDetail,stokSnapshot:stokSnap,katPengeluaran:katPenH,jumlahTransaksi:penjualanDetail.length,jumlahDO:doDetail.length}
};
setData(d=>({...d,tutupBuku:[rec,...(d.tutupBuku||[])]}));
toast("✓ Tutup buku harian tersimpan!");
}
function saveBulanan(){
var rec={id:uid(),bulan:bln,tanggal:bln+"-28",omzet:omzetB,labaKotor:marginB,totalOut:totalOutB,labaBersih:labaBersihB,assetValue,piutangA,nilaiStokA,pinjamanA,cashLaci:Number(cashLaci)||0,rekBSI:Number(rekBSI)||0,rekBCA:Number(rekBCA)||0};
setData(d=>({...d,tutupBuku:[rec,...(d.tutupBuku||[])]}));
toast("✓ Tutup buku bulanan tersimpan!");
}

function AssetSection(){return <Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>🏦 Komponen Asset</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10,marginBottom:10}}>
<Inp label="Cash di Laci" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Rekening BSI" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Rekening BCA" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
{[["Piutang Aktif (auto)",fR(piutangA),C.olt],["Nilai Stok Isi (auto)",fR(nilaiStokA),C.blt],["Pinjaman Karyawan (auto)",fR(pinjamanA),C.gl2],["Tabung Kosong",kosong55+"×5.5 / "+kosong12+"×12 / "+kosong50+"×50 (qty)",C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
<div style={{display:"flex",justifyContent:"space-between",padding:"12px 14px",background:C.nav,borderRadius:8,border:"2px solid "+C.olt}}>
<b style={{fontSize:14,color:C.wht}}>TOTAL ASSET VALUE</b>
<div style={{textAlign:"right"}}>
<b style={{fontSize:20,color:C.olt}}>{fR(assetValue)}</b>
{deltaAsset!==null&&<div style={{fontSize:11,color:deltaAsset>=0?C.glt:C.rlt,marginTop:2}}>{deltaAsset>=0?"▲":"▼"} {fR(Math.abs(deltaAsset))} vs sebelumnya</div>}
</div>
</div>
</Card>;}

return <div>
<STitle icon="📒" children="Tutup Buku"/>
<div style={{display:"flex",gap:6,marginBottom:14}}>{[["harian","📅 Harian"],["bulanan","📆 Bulanan"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"8px 18px",fontWeight:700,fontSize:13,cursor:"pointer"}}>{x[1]}</button>)}</div>

{tab==="harian"&&<div id="_tb_hari">
<Card><Inp label="Tanggal" type="date" value={tgl} onChange={setTgl} style={{maxWidth:220,marginBottom:0}}/></Card>

{/* CASH FLOW / OMSET */}
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>💰 CASH FLOW / OMSET — {fDs(tgl)}</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[
["Cash di Laci",Number(cashLaci)||0,C.wht,false],
["Bank BSI",Number(rekBSI)||0,"#9CA3AF",false],
["Bank BCA",Number(rekBCA)||0,"#9CA3AF",false],
["Total Cash + Bank",cashTotal,C.wht,true],
["",null,null,false],
["Stok Isi Gudang (HPP)",nilaiStokA,C.gl2,false],
["DO Gantung",nilaiDOGantung,nilaiDOGantung>0?"#F59E0B":C.gl2,false],
["BON Konsumen",piutangA,C.gl2,false],
["Pinjaman Karyawan",pinjamanA,C.gl2,false],
["PIUTANG & MODAL BERJALAN",piutangModal,C.wht,true],
["",null,null,false],
["TOTAL CASH FLOW",cashFlowOmset,cashFlowOmset>=0?C.glt:C.rlt,true],
].map((x,i)=>x[1]===null&&x[0]===""?<div key={i} style={{height:8,background:"transparent"}}/> :<div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}>
<span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span>
<span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1]||0)}</span>
</div>)}
</div>
{/* Pecah Kas Input */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📥 Input Cash Fisik:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:8,marginBottom:10}}>
<Inp label="Cash di Laci (Rp)" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Saldo Bank BSI (Rp)" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Saldo Bank BCA (Rp)" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
{/* Pecah Uang */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>🪙 Pecahan Kas Fisik (opsional):</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{DENOMS.map(d=><div key={d} style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"5px 11px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}><span style={{color:C.wht,fontSize:13,fontWeight:600}}>{fR(d)}</span><input type="number" value={pecah[d]} placeholder="0" onChange={e=>setPecah(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.wht,fontSize:12,outline:"none",width:74}}/><span style={{color:Number(pecah[d]||0)>0?C.olt:C.gl2,fontWeight:700,fontSize:12}}>{Number(pecah[d]||0)>0?fR(Number(pecah[d])*d):"-"}</span></div>)}
<div style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"9px 11px",background:C.nav,borderTop:"2px solid "+C.bdr}}><b style={{color:C.wht}}>Total Tunai</b><span/><b style={{color:C.glt}}>{fR(totalPecah)}</b></div>
</div>
</Card>

{/* P&L HARI INI */}
<Card style={{border:"1px solid "+C.glt}}>
<div style={{fontWeight:700,color:C.glt,marginBottom:12,fontSize:13}}>📊 P&L Hari Ini</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[["Omzet",omzetH,C.wht,false],["HPP / Modal",hppH,C.gl2,false],["Laba Kotor",marginH,C.blt,true],["Pengeluaran Operasional",-totalOutH,C.rlt,false],["Pemasukan Lainnya",Number(pemasukanLain)||0,C.olt,false],["LABA BERSIH",labaBersihH+(Number(pemasukanLain)||0),(labaBersihH+(Number(pemasukanLain)||0))>=0?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span><span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
<Inp label="Pemasukan Lainnya (topup saham dll)" type="number" value={pemasukanLain} onChange={setPemasukanLain} placeholder="0"/>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:10}}>
{[["Cash",cashInH,C.glt],["Transfer",tfInH,C.blt],["BON",bonInH,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
</Card>

{/* VERIFIKASI SELISIH */}
<Card style={{border:"1px solid "+(Math.abs(deltaSelisih)<1000?C.glt:C.olt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>✅ Verifikasi Cash Flow</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden"}}>
{[["Total Cash Flow Kemarin",cashFlowKemarin,C.gl2,false],["Laba Hari Ini",labaBersihH+(Number(pemasukanLain)||0),C.glt,false],["Total Cash Flow Hari Ini",cashFlowOmset,C.blt,true],["SELISIH",deltaSelisih,Math.abs(deltaSelisih)<1000?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span><span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
{Math.abs(deltaSelisih)<1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.grn,borderRadius:6,fontSize:12,fontWeight:700,color:"white"}}>✅ Selisih = 0. Cash flow balance!</div>}
{Math.abs(deltaSelisih)>=1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.rdk,borderRadius:6,fontSize:12,color:"white"}}>⚠️ Ada selisih {fR(Math.abs(deltaSelisih))}. Periksa input cash atau ada transaksi yang terlewat.</div>}
</Card>

{/* TOTAL ASSET */}
<Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>🏦 TOTAL ASSET</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[["Total Cash Flow",cashFlowOmset,C.blt,false],["Asset Tabung Milik PT",assetTabungMilikPT,C.gl2,false],["Asset Armada",assetArmada,C.gl2,false],["TOTAL ASSET",assetValue,C.olt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"12px 14px":"8px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?14:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?800:400}}>{x[0]}</span><span style={{fontSize:x[3]?18:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
</Card>

{/* REKAP TABUNG */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 Mutasi Stok Hari Ini</div>
<TabelStokHarian data={data} tgl={tgl}/>
</Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Rekap Tabung</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.nav}}>
{["Keterangan","12 kg","5,5 kg","50 kg"].map(h=><th key={h} style={{padding:"8px 10px",color:C.gl2,fontWeight:700,textAlign:h==="Keterangan"?"left":"center",borderBottom:"2px solid "+C.bdr}}>{h}</th>)}
</tr></thead>
<tbody>
{[
["Di Gudang (Isi)",(data.stock||{})["12 kg"]||0,(data.stock||{})["5.5 kg"]||0,(data.stock||{})["50 kg"]||0,C.glt],
["Kosong di Gudang",kosong12,kosong55,kosong50,C.gl2],
["Titip ke Konsumen",getTitipTotal(data.titipList,"12 kg"),getTitipTotal(data.titipList,"5.5 kg"),getTitipTotal(data.titipList,"50 kg"),C.blt],
["Total Keseluruhan Milik PT",((data.totalAsset||{})["12 kg"]||0),((data.totalAsset||{})["5.5 kg"]||0),((data.totalAsset||{})["50 kg"]||0),C.olt,true],
["Titipan Pihak Lain di PT",Math.max(0,titipLuarBal["12 kg"]||0),Math.max(0,titipLuarBal["5.5 kg"]||0),Math.max(0,titipLuarBal["50 kg"]||0),"#6B7280"],
["MILIK PT HOE TRANGSA",Math.max(0,((data.totalAsset||{})["12 kg"]||0)-(titipLuarBal["12 kg"]||0)),Math.max(0,((data.totalAsset||{})["5.5 kg"]||0)-(titipLuarBal["5.5 kg"]||0)),Math.max(0,((data.totalAsset||{})["50 kg"]||0)-(titipLuarBal["50 kg"]||0)),C.wht,true],
].map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:r[5]?C.nav:"transparent"}}>
<td style={{padding:"7px 10px",color:r[5]?C.wht:C.gl2,fontWeight:r[5]?700:400}}>{r[0]}</td>
{[r[1],r[2],r[3]].map((v,j)=><td key={j} style={{padding:"7px 10px",textAlign:"center",color:r[4],fontWeight:r[5]?800:600,fontSize:r[5]?14:12}}>{v}</td>)}
</tr>)}
</tbody>
</table>
</div>
</Card>

<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
<Btn color="green" onClick={saveHarian}>💾 Simpan Tutup Buku</Btn>
<button onClick={()=>doPrint("_tb_hari")} style={{background:C.blu,color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak</button>
<span style={{fontSize:11,color:"#888",fontStyle:"italic",alignSelf:"center"}}>Pilih "Save as PDF" di dialog cetak</span>
</div>
</div>}

{tab==="bulanan"&&<div id="_tb_bln">
<Card><MonthPicker label="Pilih Bulan" value={bln} onChange={setBln}/></Card>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>📊 P&L Bulanan — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
{/* Summary cards */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8,marginBottom:14}}>
{[["Total Omzet",omzetB,C.wht],["HPP/Modal (DO)",hppBln,C.olt],["Margin Kotor",marginB,C.blt],["Pengeluaran Ops",totalOutB,C.rlt],["Laba Bersih",labaBersihB,labaBersihB>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
{/* Grafik 1: Omzet & HPP per hari */}
{chartDataB.some(d=>d.omzet>0)&&<>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📈 Omzet & Modal HPP per Hari</div>
<ResponsiveContainer width="100%" height={180}><AreaChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="omzet" stroke={C.blt} fill={C.blt} fillOpacity={0.25} name="Omzet"/>
<Area type="monotone" dataKey="hpp" stroke={C.olt} fill={C.olt} fillOpacity={0.2} name="HPP/Modal"/>
</AreaChart></ResponsiveContainer>
{/* Grafik 2: Margin & Pengeluaran */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,marginTop:14}}>📊 Margin Kotor & Pengeluaran per Hari</div>
<ResponsiveContainer width="100%" height={160}><BarChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="marginKotor" fill={C.blt} name="Margin Kotor" opacity={0.8}/>
<Bar dataKey="pengeluaran" fill={C.rlt} name="Pengeluaran" opacity={0.8}/>
</BarChart></ResponsiveContainer>
{/* Grafik 3: Cash Flow Kumulatif */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,marginTop:14}}>💰 Cash Flow Kumulatif (Laba Bersih per Hari)</div>
<ResponsiveContainer width="100%" height={160}><AreaChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="cashFlow" stroke={C.glt} fill={C.glt} fillOpacity={0.3} name="Cash Flow Kumulatif"/>
<Area type="monotone" dataKey="labaBersih" stroke="#F59E0B" fill="#F59E0B" fillOpacity={0.15} name="Laba Bersih Harian"/>
</AreaChart></ResponsiveContainer>
</>}
</Card>
{top5.some(x=>x.omzet>0)&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>🏆 Top 5 Hari (Omzet)</div><RTbl headers={["Hari","Omzet","Laba Bersih"]} rows={top5.map(x=>[bln+"-"+x.hari,<b style={{color:C.blt}}>{fR(x.omzet)}</b>,<b style={{color:x.labaBersih>=0?C.glt:C.rlt}}>{fR(x.labaBersih)}</b>])}/></Card>}
{katPenArr.length>0&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>💸 Pengeluaran per Kategori</div><RTbl headers={["Kategori","Total","% dari Pengeluaran"]} rows={katPenArr.map(([k,v])=>[k,<b style={{color:C.rlt}}>{fR(v)}</b>,(totalOutB>0?(v/totalOutB*100).toFixed(1):0)+"%"])}/></Card>}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💰 Input Cash untuk Simpan Bulanan</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:8,marginBottom:10}}>
<Inp label="Cash di Laci (Rp)" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Saldo Bank BSI (Rp)" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Saldo Bank BCA (Rp)" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
</Card>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
<Btn color="green" onClick={saveBulanan}>💾 Simpan</Btn>
<button onClick={()=>doPrint("_tb_bln")} style={{background:C.blu,color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak</button>
<span style={{fontSize:11,color:"#888",fontStyle:"italic",alignSelf:"center"}}>Pilih "Save as PDF" di dialog cetak</span>
<button onClick={()=>{var wb=XLSX.utils.book_new();var s=[["Bulan",bln],["Omzet",omzetB],["Laba Kotor",marginB],["Total Pengeluaran",totalOutB],["Laba Bersih",labaBersihB],["Asset Value",assetValue]];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s),"Tutup Buku");XLSX.writeFile(wb,"TutupBuku_"+bln+".xlsx");toast("✓ Excel didownload!");}} style={{background:"#15803D",color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>📥 Excel</button>
</div>
</div>}

{(data.tutupBuku||[]).length>0&&<Card style={{marginTop:16}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📜 Riwayat Tutup Buku</div>
<RTbl headers={["Tanggal/Bulan","Transaksi","Omzet","Laba Bersih","Asset","Aksi"]} rows={(data.tutupBuku||[]).slice(0,30).map(r=>[
  <div><div style={{fontWeight:700,color:C.wht}}>{r.bulan?BULAN_ID[Number(r.bulan.split("-")[1])-1]+" "+r.bulan.split("-")[0]:fDs(r.tanggal)}</div><div style={{fontSize:10,color:C.gl2}}>{r.bulan?"Bulanan":"Harian"}</div></div>,
  <span style={{color:C.blt,fontWeight:700}}>{r.detail?.jumlahTransaksi||"-"} inv{r.detail?.jumlahDO?", "+r.detail.jumlahDO+" DO":""}</span>,
  fR(r.omzet||0),
  <b style={{color:(r.labaBersih||0)>=0?C.glt:C.rlt}}>{fR(r.labaBersih||0)}</b>,
  fR(r.assetValue||0),
  <button onClick={()=>setViewTB(r)} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 10px",color:C.blt,cursor:"pointer",fontSize:12,fontWeight:700}}>👁️ Lihat</button>
])}/>
</Card>}

{/* Modal Detail Tutup Buku */}
{viewTB&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:9000,overflowY:"auto",padding:16}}>
<div style={{maxWidth:680,margin:"0 auto",background:C.card,borderRadius:12,border:"1px solid "+C.bdr,padding:20}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
<div><div style={{fontSize:16,fontWeight:800,color:C.wht}}>📒 Laporan Tutup Buku</div><div style={{fontSize:12,color:C.gl2}}>{viewTB.bulan?BULAN_ID[Number(viewTB.bulan.split("-")[1])-1]+" "+viewTB.bulan.split("-")[0]:fDs(viewTB.tanggal)}</div></div>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"6px 14px",color:C.wht,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>

{/* Ringkasan P&L */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.blt,fontSize:12}}>📊 Ringkasan P&L</div>
{[["Omzet",viewTB.omzet,C.wht],["HPP / Modal",viewTB.hpp,C.gl2],["Laba Kotor",viewTB.labaKotor,C.blt],["Total Pengeluaran",viewTB.totalOut,C.rlt],["LABA BERSIH",viewTB.labaBersih,viewTB.labaBersih>=0?C.glt:C.rlt]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 14px",borderBottom:"1px solid "+C.bdr,background:i===4?C.nav:"transparent"}}><span style={{color:i===4?C.wht:C.gl2,fontWeight:i===4?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:700,fontSize:i===4?15:13}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* Komposisi Bayar */}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
{[["Cash",viewTB.cashIn,C.glt],["Transfer",viewTB.tfIn,C.blt],["BON",viewTB.bonIn,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:14,fontWeight:800,color:x[2]}}>{fR(x[1]||0)}</div></div>)}
</div>

{/* DO Masuk */}
{viewTB.detail?.doMasuk?.length>0&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>🚚 DO Masuk ({viewTB.detail.doMasuk.length} DO)</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden"}}>
{viewTB.detail.doMasuk.map((d,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}><span style={{color:C.gl2}}>{d.trip} — {d.sppbe}</span><span style={{color:C.wht,fontWeight:700}}>{d.ukuran} × {d.qty} tab · HPP {fR(d.totalHPP||0)}</span></div>)}
</div>
</div>}

{/* Pengeluaran per kategori */}
{viewTB.detail?.katPengeluaran&&Object.keys(viewTB.detail.katPengeluaran).length>0&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>💸 Pengeluaran per Kategori</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden"}}>
{Object.entries(viewTB.detail.katPengeluaran).sort((a,b)=>b[1]-a[1]).map(([k,v],i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}><span style={{color:C.gl2}}>{k}</span><span style={{color:C.rlt,fontWeight:700}}>{fR(v)}</span></div>)}
</div>
</div>}

{/* Stok Snapshot */}
{viewTB.detail?.stokSnapshot&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>📦 Snapshot Stok Saat Tutup Buku</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
{SIZES.map(s=>{var sn=viewTB.detail.stokSnapshot[s]||{};return <div key={s} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.wht,marginBottom:4}}>{s}</div>
{[["Isi",sn.isi,C.glt],["Titip",sn.titip,C.blt],["Kosong",sn.kosong,C.gl2]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",fontSize:11}}><span style={{color:C.gl2}}>{x[0]}</span><span style={{color:x[2],fontWeight:700}}>{x[1]||0}</span></div>)}
</div>;})}
</div>
</div>}

{/* Daftar Penjualan */}
{viewTB.detail?.penjualan?.length>0&&<div>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>🧾 Transaksi Penjualan ({viewTB.detail.penjualan.length} invoice)</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",maxHeight:200,overflowY:"auto"}}>
{viewTB.detail.penjualan.map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}>
<div><div style={{color:C.blt,fontWeight:700}}>{p.noInv}</div><div style={{color:C.gl2,fontSize:11}}>{p.konsumen}</div></div>
<div style={{textAlign:"right"}}><div style={{color:C.wht,fontWeight:700}}>{fR(p.total)}</div><div style={{fontSize:10}}><Bdg color={p.bayar==="bon"?"red":p.bayar==="transfer"?"blue":"green"}>{p.bayar}</Bdg></div></div>
</div>)}
</div>
</div>}

<div style={{marginTop:14,display:"flex",gap:8}}>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 18px",color:C.wht,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>
</div>}
</div>;
}

// ─── LAPORAN v4 (FilterTbl di semua tab, 2 tab baru) ──────────────────────────
function LaporanMod({data,toast}){
var C=useTheme();
var[mode,setMode]=useState("bulanan");var[bln,setBln]=useState(toMonth());var[tgl,setTgl]=useState(toDay());var[tab,setTab]=useState("ringkasan");
var penjAll=data.penjualan||[];
var penjFilt=mode==="bulanan"?penjAll.filter(p=>(p.tanggal||"").startsWith(bln)):penjAll.filter(p=>p.tanggal===tgl);
var penFilt=mode==="bulanan"?(data.pengeluaran||[]).filter(p=>(p.tanggal||"").startsWith(bln)):(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var omzet=penjFilt.reduce((a,p)=>a+(p.total||0),0);
var margin=penjFilt.reduce((a,p)=>a+(p.margin||0),0);
var pengeluaran=penFilt.reduce((a,p)=>a+Number(p.nominal||0),0);
var labaBersih=margin-pengeluaran;
var cash=penjFilt.filter(p=>p.bayar==="cash").reduce((a,p)=>a+p.total,0);
var tf=penjFilt.filter(p=>p.bayar==="transfer").reduce((a,p)=>a+p.total,0);
var bon=penjFilt.filter(p=>p.bayar==="bon").reduce((a,p)=>a+p.total,0);

// Per sales
var salesMap={};penjFilt.forEach(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var key=p.salesId||"?";if(!salesMap[key])salesMap[key]={id:key,nama:emp?.nama||"Tanpa Sales",omzet:0,margin:0,trx:0,q55:0,q12:0,q50:0};salesMap[key].omzet+=p.total||0;salesMap[key].margin+=p.margin||0;salesMap[key].trx++;(p.items||[]).forEach(it=>{if(it.ukuran==="5.5 kg")salesMap[key].q55+=Number(it.qty||0);else if(it.ukuran==="12 kg")salesMap[key].q12+=Number(it.qty||0);else if(it.ukuran==="50 kg")salesMap[key].q50+=Number(it.qty||0);});});
var salesArr=Object.values(salesMap);

// Per kategori
var katMap={};penjFilt.forEach(p=>{var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);var k=plg?.kategori||"Tanpa Kategori";if(!katMap[k])katMap[k]={kategori:k,omzet:0,trx:0,unik:new Set()};katMap[k].omzet+=p.total||0;katMap[k].trx++;katMap[k].unik.add(p.konsumen);});
var katArr=Object.values(katMap).map(x=>({...x,unik:x.unik.size}));

// Per produk
var prodArr=SIZES.map(s=>{var qty=0,omz=0;penjFilt.forEach(p=>(p.items||[]).forEach(it=>{if(it.ukuran===s){qty+=Number(it.qty||0);omz+=Number(it.qty||0)*Number(it.price||0);}}));return{ukuran:s,qty,omzet:omz};});

// Per pelanggan
var plgMap={};penjFilt.forEach(p=>{var k=p.konsumenId||p.konsumen;if(!plgMap[k])plgMap[k]={id:k,nama:p.konsumen,regNo:(data.pelanggan||[]).find(x=>x.id===p.konsumenId)?.regNo||"-",kategori:(data.pelanggan||[]).find(x=>x.id===p.konsumenId)?.kategori||"-",omzet:0,trx:0};plgMap[k].omzet+=p.total||0;plgMap[k].trx++;});
var plgArr=Object.values(plgMap);

// Per sales-kategori matrix (compact)
var skMap={};penjFilt.forEach(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var sn=emp?.nama||"?";var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);var kat=plg?.kategori||"Tanpa Kategori";var key=sn+"||"+kat;if(!skMap[key])skMap[key]={sales:sn,kategori:kat,omzet:0,trx:0};skMap[key].omzet+=p.total||0;skMap[key].trx++;});
var skArr=Object.values(skMap);

// Detail flat
var detailRows=penjFilt.map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);return{...p,salesNama:emp?.nama||"-",detail:(p.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ")};});

// Chart (hanya bulanan)
var chartData=[];
if(mode==="bulanan"){
  var dim2=daysInMonth(bln);
  var cfKumul=0;
  var doBlnLap=(data.doList||[]).filter(e=>(e.tanggal||"").startsWith(bln)&&(e.status||"diterima")==="diterima");
  for(var d=1;d<=dim2;d++){
    var ds=bln+"-"+String(d).padStart(2,"0");
    var pp=penjAll.filter(x=>x.tanggal===ds);
    var oz=pp.reduce((a,x)=>a+(x.total||0),0);
    var mg=pp.reduce((a,x)=>a+(x.margin||0),0);
    var pn=(data.pengeluaran||[]).filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.nominal||0),0);
    var hpp=doBlnLap.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0);
    var lb=mg-pn;
    cfKumul+=lb;
    chartData.push({tgl:String(d),omzet:oz,hpp,marginKotor:mg,pengeluaran:pn,labaBersih:lb,cashFlow:cfKumul});
  }
}

function exportExcel(){
var wb=XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Tgl","No.Inv","Konsumen","Sales","Total","Margin","Bayar"],...penjFilt.map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);return[p.tanggal,p.noInv||"",p.konsumen,emp?.nama||"",p.total,p.margin,p.bayar];})]),"Penjualan");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Sales","Omzet","Margin","Trx","5.5kg","12kg","50kg"],...salesArr.map(s=>[s.nama,s.omzet,s.margin,s.trx,s.q55,s.q12,s.q50])]),"Per Sales");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Kategori","Omzet","Unik","Trx"],...katArr.map(k=>[k.kategori,k.omzet,k.unik,k.trx])]),"Per Kategori");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Pelanggan","Reg","Kategori","Omzet","Trx"],...plgArr.map(p=>[p.nama,p.regNo,p.kategori,p.omzet,p.trx])]),"Per Pelanggan");
XLSX.writeFile(wb,"Laporan_HTS_"+(mode==="bulanan"?bln:tgl)+".xlsx");
toast("✓ Excel didownload!");
}

var salesCols=[{key:"nama",label:"Sales",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"margin",label:"Margin",render:r=><span style={{color:C.glt}}>{fR(r.margin)}</span>,sortVal:r=>r.margin,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false},{key:"q55",label:"5.5kg",sortVal:r=>r.q55,filterable:false},{key:"q12",label:"12kg",sortVal:r=>r.q12,filterable:false},{key:"q50",label:"50kg",sortVal:r=>r.q50,filterable:false}];
var katCols=[{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"unik",label:"Plg Unik",sortVal:r=>r.unik,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false},{key:"pct",label:"% Omzet",render:r=><span style={{color:C.gl2}}>{omzet>0?(r.omzet/omzet*100).toFixed(1):0}%</span>,filterable:false}];
var plgCols=[{key:"nama",label:"Pelanggan",filterable:true},{key:"regNo",label:"Reg",filterable:true},{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false}];
var skCols=[{key:"sales",label:"Sales",filterable:true},{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false}];
var detCols=[{key:"tanggal",label:"Tgl",render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},{key:"noInv",label:"Invoice",render:r=><span style={{fontSize:11,color:C.blt}}>{r.noInv||"-"}</span>,filterable:true},{key:"konsumen",label:"Konsumen",filterable:true},{key:"salesNama",label:"Sales",filterable:true},{key:"detail",label:"Produk",filterable:true},{key:"total",label:"Total",render:r=><b style={{color:C.wht}}>{fR(r.total)}</b>,sortVal:r=>r.total,filterable:false},{key:"bayar",label:"Bayar",render:r=>r.bayar==="bon"?<Bdg color="red">BON</Bdg>:r.bayar==="transfer"?<Bdg color="blue">TF</Bdg>:<Bdg color="green">Cash</Bdg>,filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"}]}];

return <div>
<STitle icon="📊" children="Laporan"/>
<Card>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>{[["bulanan","📆 Bulanan"],["harian","📅 Harian"]].map(x=><button key={x[0]} onClick={()=>setMode(x[0])} style={{background:mode===x[0]?C.blu:C.nav,color:mode===x[0]?"white":C.wht,border:"1px solid "+(mode===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}</div>
{mode==="bulanan"?<MonthPicker value={bln} onChange={setBln} label=""/>:<Inp type="date" value={tgl} onChange={setTgl} label=""/>}
<Btn sm color="green" onClick={exportExcel}>📥 Export Excel</Btn>
</Card>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:14}}>
{[["Omzet",omzet,C.wht,"📈"],["Laba Kotor",margin,C.blt,"💹"],["Pengeluaran",pengeluaran,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"],["Transaksi",penjFilt.length+" trx",C.gl2,"🧾"]].map(x=><SC key={x[0]} label={x[0]} value={typeof x[1]==="number"?fR(x[1]):x[1]} icon={x[3]} color={x[2]}/>)}
</div>
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>{[["ringkasan","📊 Ringkasan"],["grafik","📈 Grafik"],["sales","👤 Per Sales"],["kategori","🏷️ Per Kategori"],["produk","📦 Per Produk"],["pelanggan","👥 Per Pelanggan"],["matrix","📋 Sales×Kategori"],["detail","🔍 Detail"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"6px 11px",fontWeight:700,fontSize:11,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="ringkasan"&&<>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💰 Komposisi Pembayaran</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>{[["Cash",cash,C.glt],["Transfer",tf,C.blt],["BON",bon,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr,textAlign:"center"}}><div style={{fontSize:11,color:C.gl2}}>{x[0]}</div><div style={{fontSize:14,fontWeight:900,color:x[2]}}>{fR(x[1])}</div><div style={{fontSize:10,color:C.gl2,marginTop:2}}>{omzet>0?(x[1]/omzet*100).toFixed(1):0}%</div></div>)}</div></Card>

</>}
{tab==="sales"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👤 Per Sales</div><FilterTbl columns={salesCols} data={salesArr} empty="Tidak ada data"/></Card>}
{tab==="kategori"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🏷️ Per Kategori Pelanggan</div><FilterTbl columns={katCols} data={katArr} empty="Tidak ada data"/>{katArr.length>0&&<div style={{marginTop:14}}><ResponsiveContainer width="100%" height={220}><PieChart><Pie data={katArr.slice(0,8)} dataKey="omzet" nameKey="kategori" cx="50%" cy="50%" outerRadius={80} label={x=>x.kategori}>{katArr.slice(0,8).map((e,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}</Pie><Tooltip formatter={v=>fR(v)} contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht}}/></PieChart></ResponsiveContainer></div>}</Card>}
{tab==="produk"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Per Produk</div><RTbl headers={["Ukuran","Qty","Omzet","% Omzet"]} rows={prodArr.map(p=>[<b style={{color:C.wht}}>{p.ukuran}</b>,<b style={{color:C.glt}}>{p.qty} tab</b>,<b style={{color:C.blt}}>{fR(p.omzet)}</b>,(omzet>0?(p.omzet/omzet*100).toFixed(1):0)+"%"])}/></Card>}
{tab==="pelanggan"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👥 Ranking Pelanggan per Omzet</div><FilterTbl columns={plgCols} data={plgArr} empty="Tidak ada data"/></Card>}
{tab==="matrix"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Sales × Kategori</div><FilterTbl columns={skCols} data={skArr} empty="Tidak ada data"/></Card>}
{tab==="detail"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔍 Detail Penjualan ({penjFilt.length})</div><FilterTbl columns={detCols} data={detailRows} empty="Tidak ada data" maxRows={300}/></Card>}

{tab==="grafik"&&<div>
{mode!=="bulanan"?<Card><div style={{color:C.gl2,fontSize:13,textAlign:"center",padding:20}}>📈 Grafik hanya tersedia untuk mode Bulanan</div></Card>:
!chartData.some(d=>d.omzet>0)?<Card><div style={{color:C.gl2,fontSize:13,textAlign:"center",padding:20}}>Belum ada data penjualan bulan ini</div></Card>:<>

{/* Ringkasan bulanan */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8,marginBottom:14}}>
{[["Total Omzet",omzet,C.wht,"📈"],["HPP/Modal DO",doBlnLap.reduce((a,d)=>a+Number(d.totalHPP||0),0),C.olt,"🏭"],["Margin Kotor",margin,C.blt,"💹"],["Pengeluaran Ops",pengeluaran,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"],["Cash Flow Akhir",chartData[chartData.length-1]?.cashFlow||0,(chartData[chartData.length-1]?.cashFlow||0)>=0?C.glt:C.rlt,"💰"]].map(x=><div key={x[0]} style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2,marginBottom:2}}>{x[3]} {x[0]}</div><div style={{fontSize:13,fontWeight:800,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>

{/* Grafik 1: Omzet & HPP */}
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>📈 1. Omzet Penjualan & Modal HPP per Hari</div>
<ResponsiveContainer width="100%" height={200}><AreaChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<Legend wrapperStyle={{fontSize:11,color:C.gl2}}/>
<Area type="monotone" dataKey="omzet" stroke="#3B82F6" fill="#3B82F6" fillOpacity={0.25} name="Omzet"/>
<Area type="monotone" dataKey="hpp" stroke="#F59E0B" fill="#F59E0B" fillOpacity={0.2} name="HPP/Modal DO"/>
</AreaChart></ResponsiveContainer></Card>

{/* Grafik 2: Margin vs Pengeluaran */}
<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>💹 2. Margin Kotor vs Pengeluaran Operasional per Hari</div>
<ResponsiveContainer width="100%" height={200}><BarChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<Legend wrapperStyle={{fontSize:11,color:C.gl2}}/>
<Bar dataKey="marginKotor" fill="#3B82F6" name="Margin Kotor" opacity={0.85}/>
<Bar dataKey="pengeluaran" fill="#EF4444" name="Pengeluaran Ops" opacity={0.85}/>
</BarChart></ResponsiveContainer></Card>

{/* Grafik 3: Laba Bersih per hari */}
<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>🏆 3. Laba Bersih per Hari (bar hijau=untung, merah=rugi)</div>
<ResponsiveContainer width="100%" height={180}><BarChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
{chartData.map((entry,index)=><Cell key={index} fill={entry.labaBersih>=0?"#22C55E":"#EF4444"}/>)}
<Bar dataKey="labaBersih" name="Laba Bersih">{chartData.map((entry,index)=><Cell key={index} fill={entry.labaBersih>=0?"#22C55E":"#EF4444"}/>)}</Bar>
</BarChart></ResponsiveContainer></Card>

{/* Grafik 4: Cash Flow Kumulatif — PALING PENTING */}
<Card style={{marginTop:10,border:"2px solid "+C.glt}}><div style={{fontWeight:700,color:C.glt,marginBottom:8,fontSize:12}}>💰 4. Cash Flow Kumulatif Bulanan (PALING PENTING)</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Akumulasi laba bersih dari hari 1 sampai akhir bulan — naik = bisnis sehat</div>
<ResponsiveContainer width="100%" height={220}><AreaChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<defs><linearGradient id="cfGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#22C55E" stopOpacity={0.4}/><stop offset="95%" stopColor="#22C55E" stopOpacity={0}/></linearGradient></defs>
<Area type="monotone" dataKey="cashFlow" stroke="#22C55E" fill="url(#cfGrad)" strokeWidth={2} name="Cash Flow Kumulatif"/>
<ReferenceLine y={0} stroke={C.rlt} strokeDasharray="4 4"/>
</AreaChart></ResponsiveContainer></Card>

{/* Grafik 5: Top 10 Konsumen */}
{plgArr.length>0&&<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>👥 5. Top 10 Pelanggan bulan ini</div>
<ResponsiveContainer width="100%" height={220}><BarChart layout="vertical" data={plgArr.slice().sort((a,b)=>b.omzet-a.omzet).slice(0,10)} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="nama" stroke={C.gl2} fontSize={9} width={100}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="omzet" fill="#3B82F6" name="Omzet" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer></Card>}

{/* Grafik 6: Per Sales */}
{salesArr.length>0&&<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>👤 6. Omzet per Sales</div>
<ResponsiveContainer width="100%" height={Math.max(160,salesArr.length*45)}><BarChart layout="vertical" data={salesArr.slice().sort((a,b)=>b.omzet-a.omzet)} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="nama" stroke={C.gl2} fontSize={9} width={100}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="omzet" fill="#8B5CF6" name="Omzet" radius={[0,4,4,0]}/>
<Bar dataKey="margin" fill="#22C55E" name="Margin" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer></Card>}

</>}
</div>}

</div>;
}

// ─── KARYAWAN & AMBILAN ───────────────────────────────────────────────────────
function KaryawanMod({data,setData,toast}){
var C=useTheme();
var blk={username:"",password:"",role:"sales_driver",nama:"",posisi:"Sales Driver",telepon:"",alamat:"",gajiPokok:"",uangMakan:"15000",uangMakanMode:"harian",aktif:true};
var[f,setF]=useState({...blk});var[edit,setEdit]=useState(null);var[delId,setDelId]=useState(null);
var[ambF,setAmbF]=useState({empId:"",nominal:"",ket:"",tanggal:toDay()});
var posOpts=["Owner/Komisaris","Manajer","Admin","Kasir/Akuntan","Sales Driver","Sales Freelance","Sales Marketing","Checker","Driver Truck SPBE","Helper"];
function save(){if(!f.nama||!f.username)return;if(edit){setData(d=>({...d,employees:(d.employees||[]).map(e=>e.id===edit.id?{...e,...f,gajiPokok:Number(f.gajiPokok||0),uangMakan:Number(f.uangMakan||15000)}:e)}));setEdit(null);}else setData(d=>({...d,employees:[{id:uid(),...f,gajiPokok:Number(f.gajiPokok||0),uangMakan:Number(f.uangMakan||15000)},...(d.employees||[])]}));setF({...blk});toast("✓ Karyawan disimpan!");}
function saveAmb(){if(!ambF.empId||!ambF.nominal)return;var emp=(data.employees||[]).find(e=>e.id===ambF.empId);setData(d=>({...d,ambilan:[{id:uid(),karyawanId:ambF.empId,karyawanNama:emp?.nama||"",nominal:Number(ambF.nominal),ket:ambF.ket,tanggal:ambF.tanggal},...(d.ambilan||[])]}));setAmbF(p=>({...p,nominal:"",ket:""}));toast("✓ Ambilan dicatat!");}
return <div>
<STitle icon="👤" children="Karyawan & Akun"/>
<Card><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><Sel label="Posisi" value={f.posisi} onChange={v=>setF(p=>({...p,posisi:v}))} opts={posOpts}/><Sel label="Role" value={f.role} onChange={v=>setF(p=>({...p,role:v}))} opts={Object.keys(ROLE_LBL).map(k=>({v:k,l:ROLE_LBL[k]}))}/><Inp label="Username" value={f.username} onChange={v=>setF(p=>({...p,username:v}))}/><Inp label="Password" type="password" value={f.password} onChange={v=>setF(p=>({...p,password:v}))}/><Inp label="Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/><Inp label="Gaji Pokok" type="number" value={f.gajiPokok} onChange={v=>setF(p=>({...p,gajiPokok:v}))}/><Inp label="Uang Makan/Hari" type="number" value={f.uangMakan} onChange={v=>setF(p=>({...p,uangMakan:v}))}/><Sel label="Mode Uang Makan" value={f.uangMakanMode} onChange={v=>setF(p=>({...p,uangMakanMode:v}))} opts={[{v:"harian",l:"💰 Harian"},{v:"akhir_bulan",l:"📅 Akhir Bulan"}]}/></div><Btn color="green" onClick={save} dis={!f.nama||!f.username}>➕ Tambah Karyawan</Btn></Card>
<Card><RTbl headers={["Nama","Posisi","Role","Status","Aksi"]} rows={(data.employees||[]).map(e=>[<div><b style={{color:C.wht}}>{e.nama}</b><div style={{fontSize:10,color:C.gl2}}>{e.telepon}</div></div>,e.posisi,<Bdg color={["admin","owner"].includes(e.role)?"red":"blue"}>{ROLE_LBL[e.role]||e.role}</Bdg>,e.aktif?<Bdg color="green">Aktif</Bdg>:<Bdg color="gray">Non-aktif</Bdg>,<div style={{display:"flex",gap:4}}><button onClick={()=>setData(d=>({...d,employees:(d.employees||[]).map(x=>x.id===e.id?{...x,aktif:!x.aktif}:x)}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.gl2,cursor:"pointer",fontSize:11}}>{e.aktif?"🔒":"🔓"}</button><ActBtns onEdit={()=>{setEdit(e);setF({...e,gajiPokok:String(e.gajiPokok||""),uangMakan:String(e.uangMakan||15000)});}} onDel={()=>setDelId(e)}/></div>])}/></Card>
<Card><div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>💸 Ambilan / Kasbon Karyawan</div><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10}}><Inp label="Tanggal" type="date" value={ambF.tanggal} onChange={v=>setAmbF(p=>({...p,tanggal:v}))}/><Sel label="Karyawan" value={ambF.empId} onChange={v=>setAmbF(p=>({...p,empId:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).map(e=>({v:e.id,l:e.nama}))]}/><Inp label="Nominal" type="number" value={ambF.nominal} onChange={v=>setAmbF(p=>({...p,nominal:v}))}/><Inp label="Keterangan" value={ambF.ket} onChange={v=>setAmbF(p=>({...p,ket:v}))}/></div><Btn color="orange" onClick={saveAmb} dis={!ambF.empId||!ambF.nominal}>💾 Simpan Ambilan</Btn></Card>
{(data.ambilan||[]).length>0&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>Riwayat Ambilan</div><RTbl headers={["Tgl","Karyawan","Nominal","Ket","Aksi"]} rows={(data.ambilan||[]).slice(0,30).map(a=>[fDs(a.tanggal),a.karyawanNama,<b style={{color:C.olt}}>{fR(a.nominal)}</b>,a.ket||"-",<button onClick={()=>setData(d=>({...d,ambilan:(d.ambilan||[]).filter(x=>x.id!==a.id)}))} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button>])}/></Card>}
{edit&&<Modal title={"Edit: "+edit.nama} onSave={save} onClose={()=>{setEdit(null);setF({...blk});}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><Sel label="Posisi" value={f.posisi} onChange={v=>setF(p=>({...p,posisi:v}))} opts={posOpts}/><Sel label="Role" value={f.role} onChange={v=>setF(p=>({...p,role:v}))} opts={Object.keys(ROLE_LBL).map(k=>({v:k,l:ROLE_LBL[k]}))}/><Inp label="Username" value={f.username} onChange={v=>setF(p=>({...p,username:v}))}/><Inp label="Password" type="password" value={f.password} onChange={v=>setF(p=>({...p,password:v}))}/><Inp label="Gaji Pokok" type="number" value={f.gajiPokok} onChange={v=>setF(p=>({...p,gajiPokok:v}))}/><Inp label="Uang Makan" type="number" value={f.uangMakan} onChange={v=>setF(p=>({...p,uangMakan:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/></div><Sel label="Mode Uang Makan" value={f.uangMakanMode} onChange={v=>setF(p=>({...p,uangMakanMode:v}))} opts={[{v:"harian",l:"💰 Harian"},{v:"akhir_bulan",l:"📅 Akhir Bulan"}]}/></Modal>}
{delId&&<ConfirmDel msg={"Hapus \""+delId.nama+"\"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,employees:(d.employees||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// ─── ABSENSI ──────────────────────────────────────────────────────────────────
function AbsensiMod({data,setData,toast}){
var C=useTheme();
var[f,setF]=useState({tanggal:toDay(),karyawanId:"",status:"Hadir",ket:""});var[viewBln,setViewBln]=useState(toMonth());
function save(){if(!f.karyawanId)return;var emp=(data.employees||[]).find(e=>e.id===f.karyawanId);var exists=(data.absensi||[]).find(a=>a.karyawanId===f.karyawanId&&a.tanggal===f.tanggal);var rec={id:exists?.id||uid(),karyawanId:f.karyawanId,karyawanNama:emp?.nama||"",tanggal:f.tanggal,status:f.status,ket:f.ket};setData(d=>({...d,absensi:exists?(d.absensi||[]).map(a=>a.id===exists.id?rec:a):[rec,...(d.absensi||[])]}));toast("✓ Absensi dicatat!");}
var absBln=(data.absensi||[]).filter(a=>(a.tanggal||"").startsWith(viewBln));
var rekapMap={};(data.employees||[]).filter(e=>e.aktif).forEach(e=>{var ea=absBln.filter(a=>a.karyawanId===e.id);var counts={};ABSENSI_STATUS.forEach(s=>{counts[s]=ea.filter(a=>a.status===s).length;});rekapMap[e.id]={nama:e.nama,counts};});
return <div>
<STitle icon="📅" children="Absensi"/>
<Card><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10}}><Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/><Sel label="Karyawan" value={f.karyawanId} onChange={v=>setF(p=>({...p,karyawanId:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).filter(e=>e.aktif).map(e=>({v:e.id,l:e.nama}))]}/><div><label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Status</label><div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{ABSENSI_STATUS.map(s=><button key={s} onClick={()=>setF(p=>({...p,status:s}))} style={{background:f.status===s?C.blu:C.nav,color:f.status===s?"white":C.wht,border:"1px solid "+(f.status===s?C.blt:C.bdr),borderRadius:6,padding:"5px 8px",fontSize:11,fontWeight:700,cursor:"pointer"}}>{s}</button>)}</div></div><Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/></div><Btn color="green" onClick={save} dis={!f.karyawanId}>💾 Simpan</Btn></Card>
<Card><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}><div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Rekap Bulanan</div><MonthPicker value={viewBln} onChange={setViewBln} label=""/></div>
<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr><th style={{padding:"8px 10px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"left",borderBottom:"2px solid "+C.bdr}}>Nama</th>{ABSENSI_STATUS.map(s=><th key={s} style={{padding:"8px 6px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"center",borderBottom:"2px solid "+C.bdr,fontSize:10}}>{s}</th>)}</tr></thead><tbody>{Object.keys(rekapMap).map(id=>{var r=rekapMap[id];return <tr key={id} style={{borderBottom:"1px solid "+C.bdr}}><td style={{padding:"8px 10px",color:C.wht,fontWeight:600}}>{r.nama}</td>{ABSENSI_STATUS.map(s=><td key={s} style={{padding:"7px 6px",textAlign:"center",color:r.counts[s]>0?C.wht:C.gry,fontWeight:r.counts[s]>0?700:400}}>{r.counts[s]>0?r.counts[s]:"-"}</td>)}</tr>;})}</tbody></table></div>
</Card>
</div>;
}

// ─── SETTINGS ─────────────────────────────────────────────────────────────────
function SettingsMod({data,setData,toast,theme,setTheme}){
var C=useTheme();
var[f,setF]=useState({...data.company});
var logoRef=useRef();var logoPRef=useRef();var ttdKRef=useRef();var ttdDRef=useRef();var stempelRef=useRef();
var[hP,setHP]=useState({jenis:"Isi",ukuran:"5.5 kg",harga:""});
var[mP,setMP]=useState({jenis:"Isi",ukuran:"5.5 kg",harga:"",tgl:toDay()});
function saveCompany(){setData(d=>({...d,company:{...d.company,...f}}));toast("✓ Disimpan!");}
function handleUpload(e,key,label){var file=e.target.files[0];if(!file)return;if(file.size>2*1024*1024){toast("File terlalu besar (maks 2MB)","error");return;}var reader=new FileReader();reader.onload=ev=>{var dataUrl=ev.target.result;setF(p=>({...p,[key]:dataUrl}));setData(d=>({...d,company:{...d.company,[key]:dataUrl}}));toast("✓ "+label+" diupload!");};reader.readAsDataURL(file);}
function clearUpload(key,label){setF(p=>({...p,[key]:""}));setData(d=>({...d,company:{...d.company,[key]:""}}));toast("✓ "+label+" dihapus");}
function saveHET(){if(!hP.harga)return;setData(d=>{var hp={...d.hetPrices||{}};if(!hp[hP.jenis])hp[hP.jenis]={};hp[hP.jenis][hP.ukuran]=Number(hP.harga);return{...d,hetPrices:hp};});toast("✓ HET diperbarui!");}
function saveModal(){if(!mP.harga)return;setData(d=>({...d,modalHistory:[{id:uid(),tanggal:mP.tgl,jenis:mP.jenis,ukuran:mP.ukuran,harga:Number(mP.harga),sumber:"Manual Settings"},...(d.modalHistory||[])]}));toast("✓ Modal diperbarui!");}
function clearData(){if(window.confirm("⚠️ Hapus semua data transaksi? Tidak bisa dibatalkan!")){{setData(d=>({...d,penjualan:[],bon:[],pengeluaran:[],tutupBuku:[],setoranSales:[],setoranLog:[],stockLog:[],absensi:[],ambilan:[],payrollLog:[],titipList:[],counters:{inv:{},sg:{},reg:0}}));toast("✓ Data transaksi dihapus!");}}}
function backup(){try{var blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="backup_hoetrangsa_"+toDay()+".json";a.click();toast("✓ Backup didownload!");}catch(e){toast("Gagal","error");}}
function UploadSlot({label,value,refEl,onUpload,onClear,desc}){return <div style={{display:"flex",alignItems:"center",gap:12,padding:10,background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,marginBottom:8}}>
{value?<img src={value} style={{height:60,maxWidth:120,objectFit:"contain",background:"white",borderRadius:4,padding:4}} alt={label}/>:<div style={{height:60,width:80,background:C.bg,borderRadius:6,border:"2px dashed "+C.bdr,display:"flex",alignItems:"center",justifyContent:"center",color:C.gl2,fontSize:10,textAlign:"center"}}>Belum<br/>upload</div>}
<div style={{flex:1}}><div style={{color:C.wht,fontWeight:700,fontSize:12}}>{label}</div><div style={{color:C.gl2,fontSize:10,marginBottom:6}}>{desc}</div><div style={{display:"flex",gap:6}}><input ref={refEl} type="file" accept="image/*" onChange={onUpload} style={{display:"none"}}/><Btn sm color="blue" onClick={()=>refEl.current?.click()}>📁 Upload</Btn>{value&&<Btn sm color="red" onClick={onClear}>🗑️ Hapus</Btn>}</div></div>
</div>;}
return <div>
<STitle icon="⚙️" children="Pengaturan"/>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:10,fontSize:13}}>🎨 Tema Aplikasi</div>
<div style={{display:"flex",gap:8}}>
<button onClick={()=>setTheme("light")} style={{flex:1,padding:"12px",background:theme==="light"?C.blu:C.nav,color:theme==="light"?"white":C.wht,border:"1px solid "+(theme==="light"?C.blt:C.bdr),borderRadius:8,fontWeight:700,fontSize:13,cursor:"pointer"}}>☀️ Terang Pertamina</button>
<button onClick={()=>setTheme("dark")} style={{flex:1,padding:"12px",background:theme==="dark"?C.blu:C.nav,color:theme==="dark"?"white":C.wht,border:"1px solid "+(theme==="dark"?C.blt:C.bdr),borderRadius:8,fontWeight:700,fontSize:13,cursor:"pointer"}}>🌙 Gelap Premium</button>
</div>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🖼️ Logo & Tanda Tangan Elektronik</div>
<UploadSlot label="Logo PT. Hoe Trang Sa" value={f.logo} refEl={logoRef} onUpload={e=>handleUpload(e,"logo","Logo")} onClear={()=>clearUpload("logo","Logo")} desc="Tampil di semua dokumen. Kalau belum upload, nama perusahaan ditampilkan."/>
<UploadSlot label="Logo Pertamina" value={f.logoPertamina} refEl={logoPRef} onUpload={e=>handleUpload(e,"logoPertamina","Logo Pertamina")} onClear={()=>clearUpload("logoPertamina","Logo Pertamina")} desc="Pojok kanan dokumen (opsional)"/>
<UploadSlot label="TTD Kasir" value={f.ttdKasir} refEl={ttdKRef} onUpload={e=>handleUpload(e,"ttdKasir","TTD Kasir")} onClear={()=>clearUpload("ttdKasir","TTD Kasir")} desc="Tanda tangan di Invoice"/>
<UploadSlot label="TTD Direktur" value={f.ttdDirektur} refEl={ttdDRef} onUpload={e=>handleUpload(e,"ttdDirektur","TTD Direktur")} onClear={()=>clearUpload("ttdDirektur","TTD Direktur")} desc="Tanda tangan di Kwitansi Slip Gaji"/>
<UploadSlot label="Stempel LUNAS" value={f.stempelLunas} refEl={stempelRef} onUpload={e=>handleUpload(e,"stempelLunas","Stempel LUNAS")} onClear={()=>clearUpload("stempelLunas","Stempel LUNAS")} desc="Jika kosong, pakai teks LUNAS default"/>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🏢 Data Perusahaan</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
{[["nama","Nama Perusahaan"],["slogan","Slogan"],["alamat","Alamat"],["telepon","No. HP / WA"],["telepon2","No. Telp Kantor"],["email","Email"],["website","Website"],["npwp","NPWP"],["direkturNama","Nama Direktur (slip gaji)"],["kasirNama","Nama Kasir (invoice)"]].map(([k,l])=><Inp key={k} label={l} value={f[k]||""} onChange={v=>setF(p=>({...p,[k]:v}))}/>)}
</div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan Data Perusahaan</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💳 Rekening Bank (tampil di Invoice)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}><Inp label="Nama Bank" value={f.bankNama||""} onChange={v=>setF(p=>({...p,bankNama:v}))}/><Inp label="Atas Nama" value={f.bankAtasNama||""} onChange={v=>setF(p=>({...p,bankAtasNama:v}))}/><Inp label="No. Rekening" value={f.bankRekening||""} onChange={v=>setF(p=>({...p,bankRekening:v}))}/></div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan Rekening</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💲 Update HET</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10}}><Sel label="Jenis" value={hP.jenis} onChange={v=>setHP(p=>({...p,jenis:v}))} opts={JENIS}/><Sel label="Ukuran" value={hP.ukuran} onChange={v=>setHP(p=>({...p,ukuran:v}))} opts={SIZES}/><Inp label={"HET ("+fR(getHET(data,hP.ukuran,hP.jenis))+")"} type="number" value={hP.harga} onChange={v=>setHP(p=>({...p,harga:v}))}/></div>
<Btn color="orange" onClick={saveHET} dis={!hP.harga}>💾 Update HET</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📦 Update Modal/HPP Manual</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10}}><Sel label="Jenis" value={mP.jenis} onChange={v=>setMP(p=>({...p,jenis:v}))} opts={JENIS}/><Sel label="Ukuran" value={mP.ukuran} onChange={v=>setMP(p=>({...p,ukuran:v}))} opts={SIZES}/><Inp label={"Modal ("+fR(getModal(data,mP.ukuran,mP.jenis))+")"} type="number" value={mP.harga} onChange={v=>setMP(p=>({...p,harga:v}))}/><Inp label="Berlaku Dari" type="date" value={mP.tgl} onChange={v=>setMP(p=>({...p,tgl:v}))}/></div>
<Btn color="green" onClick={saveModal} dis={!mP.harga}>💾 Update Modal</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🚛 Asset & Harga Tabung</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10}}>Untuk kalkulasi Asset Tabung Milik PT di Tutup Buku:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:12}}>
<Inp label="Asset Armada (Rp)" type="number" value={f.assetArmada||""} onChange={v=>setF(p=>({...p,assetArmada:Number(v)||0}))} placeholder="Nilai kendaraan operasional"/>
<Inp label="Harga Tbg Kosong 5,5kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["5.5 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["5.5 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
<Inp label="Harga Tbg Kosong 12kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["12 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["12 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
<Inp label="Harga Tbg Kosong 50kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["50 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["50 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
</div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🏭 Kode Pertamina</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10,fontStyle:"italic"}}>🔒 Kode penebusan (permanen — tidak berubah)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10,marginBottom:12}}>
{[["soldTo","Sold To"],["shipToKCR","Ship To KCR"],["shipToMGL","Ship To MGL"]].map(([k,l])=><Inp key={k} label={l} value={f[k]||""} onChange={v=>setF(p=>({...p,[k]:v}))}/>)}
</div>
<div style={{fontSize:11,color:C.olt,marginBottom:10,fontStyle:"italic"}}>🔄 Kode SA bulanan (update setiap bulan terima SA baru dari Pertamina)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10}}>
<Inp label="Bulan SA Aktif" value={f.saBulan||""} onChange={v=>setF(p=>({...p,saBulan:v}))} placeholder="Juni 2026"/>
<Inp label="SA 12kg — KCR" value={f.sa12KCR||""} onChange={v=>setF(p=>({...p,sa12KCR:v}))}/>
<Inp label="SA 5,5kg — KCR" value={f.sa55KCR||""} onChange={v=>setF(p=>({...p,sa55KCR:v}))}/>
<Inp label="SA 12kg — MGL" value={f.sa12MGL||""} onChange={v=>setF(p=>({...p,sa12MGL:v}))}/>
<Inp label="SA 5,5kg — MGL" value={f.sa55MGL||""} onChange={v=>setF(p=>({...p,sa55MGL:v}))}/>
</div>
<Btn color="blue" onClick={saveCompany} style={{marginTop:10}}>💾 Simpan Kode Pertamina</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💾 Backup & Data</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn color="green" onClick={backup}>📥 Export Backup JSON</Btn><Btn color="red" onClick={clearData}>🗑️ Hapus Semua Transaksi</Btn></div></Card>
</div>;
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
var ALL_TABS=[
{id:"dashboard",label:"Dashboard",icon:"🏠"},{id:"penjualan",label:"Penjualan",icon:"🧾"},
{id:"piutang",label:"Piutang",icon:"💳"},{id:"setoran",label:"Setoran",icon:"💰"},
{id:"laporan",label:"Laporan",icon:"📊"},{id:"stok",label:"Stok",icon:"📦"},
{id:"pengeluaran",label:"Pengeluaran",icon:"💸"},{id:"tutupbuku",label:"Tutup Buku",icon:"📒"},
{id:"pelanggan",label:"Pelanggan",icon:"👥"},{id:"karyawan",label:"Karyawan",icon:"👤"},
{id:"absensi",label:"Absensi",icon:"📅"},{id:"payroll",label:"Payroll",icon:"💼"},
{id:"do",label:"DO",icon:"🚚"},{id:"settings",label:"Pengaturan",icon:"⚙️"},
];
function getVisibleTabs(user){if(!user)return[];var rt=ROLE_TABS[user.role];if(rt===null||rt===undefined)return ALL_TABS;return ALL_TABS.filter(t=>rt.includes(t.id));}

export default function App(){
var[theme,setTheme]=useState(()=>{try{var s=localStorage.getItem("hts_theme");return s||"light";}catch(e){return"light";}});
var C=THEMES[theme]||THEMES.light;
useEffect(()=>{try{localStorage.setItem("hts_theme",theme);}catch(e){}},[theme]);
var[data,setData]=useState(()=>{
try{var s=localStorage.getItem("lpg_mgmt_v4");
if(s){var d=JSON.parse(s);
if(!d.totalAsset)d.totalAsset={"5.5 kg":0,"12 kg":0,"50 kg":0};
if(!d.stokKosong)d.stokKosong={"5.5 kg":0,"12 kg":0,"50 kg":0};
if(d.pelanggan)d.pelanggan=d.pelanggan.map(p=>({...p,hargaKhusus:Array.isArray(p.hargaKhusus)?p.hargaKhusus:[]}));
// Migration titipan flat format → items array
if(d.titipList)d.titipList=d.titipList.map(t=>{
  if(t.items&&t.items.length>0)return t;
  if(t.ukuran&&t.qty)return{...t,items:[{ukuran:t.ukuran,qty:Number(t.qty||0)}]};
  return t;
});
// Migration DO lama → status diterima (stok sudah masuk)
if(d.doList)d.doList=d.doList.map(x=>x.status?x:{...x,status:"diterima"});
if(!d.titipList)d.titipList=[];
if(!d.counters)d.counters={inv:{},sg:{},reg:0};
if(!d.company)d.company={...INIT.company};
if(!d.setoranLog)d.setoranLog=[];
Object.keys(INIT.company).forEach(k=>{if(d.company[k]==null)d.company[k]=INIT.company[k];});
if(!d.employees||d.employees.length<5)d.employees=DEF_EMP.slice();
d.employees=d.employees.map(e=>e.uangMakanMode?e:{...e,uangMakanMode:"harian"});
if(d.pelanggan&&d.pelanggan.length){var maxReg=d.counters.reg||0;d.pelanggan=d.pelanggan.map(p=>{if(p.regNo)return p;maxReg++;return{...p,regNo:"HTS/CST/"+String(maxReg).padStart(3,"0")};});d.counters.reg=Math.max(d.counters.reg||0,maxReg);}
return{...INIT,...d};}
// Try migrate from v3
try{var s3=localStorage.getItem("lpg_mgmt_v3");if(s3){var d3=JSON.parse(s3);return{...INIT,...d3,setoranLog:[]};}}catch(e){}
return{...INIT};}catch(e){return{...INIT};}
});
var[user,setUser]=useState(null);var[tab,setTab]=useState("dashboard");var[inv,setInv]=useState(null);var[sideOpen,setSideOpen]=useState(false);
var[syncStatus,setSyncStatus]=useState("idle");// idle|pushing|pulling|ok|error
var[syncMsg,setSyncMsg]=useState("");

async function handlePush(){
  setSyncMsg("Menyimpan ke cloud...");
  var ok=await pushAll(data,setSyncStatus);
  setSyncMsg(ok?"✅ Tersimpan ke Google Sheets!":"❌ Gagal push. Cek koneksi.");
  setTimeout(()=>setSyncMsg(""),4000);
}
async function handlePull(){
  setSyncMsg("Mengambil data dari cloud...");
  var pulled=await pullAll(setSyncStatus);
  if(pulled){
    setData(prev=>({...prev,...pulled,
      employees:pulled.employees&&pulled.employees.length>0?pulled.employees:prev.employees
    }));
    setSyncMsg("✅ Data berhasil diambil dari cloud!");
  }else{
    setSyncMsg("❌ Gagal pull. Cek koneksi.");
  }
  setTimeout(()=>setSyncMsg(""),4000);
}
var{toasts,toast}=useToast();var mobile=useMobile();
useEffect(()=>{try{localStorage.setItem("lpg_mgmt_v4",JSON.stringify(data));}catch(e){console.warn("Storage full");}},[data]);
var setDataP=useCallback(updater=>{setData(prev=>typeof updater==="function"?updater(prev):updater);},[]);
var themeToggle=()=>setTheme(t=>t==="light"?"dark":"light");
if(!user)return <ThemeCtx.Provider value={C}><LoginScreen employees={data.employees||DEF_EMP} onLogin={u=>{setUser(u);setTab("dashboard");}} themeToggle={themeToggle} theme={theme}/><Toast toasts={toasts}/>
{syncMsg&&<div style={{position:"fixed",top:60,left:"50%",transform:"translateX(-50%)",background:syncMsg.startsWith("✅")?C.grn:syncMsg.startsWith("❌")?C.rdk:C.org,color:"#fff",padding:"8px 20px",borderRadius:20,fontSize:13,fontWeight:700,zIndex:9998,boxShadow:"0 4px 16px rgba(0,0,0,.2)",whiteSpace:"nowrap"}}>{syncMsg}</div>}</ThemeCtx.Provider>;
var tabs=getVisibleTabs(user);var curTab=tabs.find(t=>t.id===tab)||tabs[0];
var bonAlerts=(data.bon||[]).filter(b=>b.status!=="lunas"&&b.deadline&&dLeft(b.deadline)!==null&&dLeft(b.deadline)<=3).length;
function renderContent(){
var t=curTab?.id;var props={data,setData:setDataP,user,toast};
if(t==="dashboard")return <Dashboard data={data} setTab={setTab} user={user}/>;
if(t==="penjualan")return <PenjualanMod {...props} setInv={setInv}/>;
if(t==="piutang")return <PiutangMod {...props} setInv={setInv}/>;
if(t==="setoran")return <SetoranMod {...props}/>;
if(t==="laporan")return <LaporanMod data={data} toast={toast}/>;
if(t==="stok")return <StokMod {...props}/>;
if(t==="pengeluaran")return <PengeluaranMod {...props}/>;
if(t==="tutupbuku")return <TutupBukuMod {...props}/>;
if(t==="pelanggan")return <PelangganMod {...props}/>;
if(t==="karyawan")return <KaryawanMod {...props}/>;
if(t==="absensi")return <AbsensiMod {...props}/>;
if(t==="payroll")return <PayrollMod {...props}/>;
if(t==="do")return <DOMod {...props}/>;
if(t==="settings")return <SettingsMod data={data} setData={setDataP} toast={toast} theme={theme} setTheme={setTheme}/>;
return null;}
return <ThemeCtx.Provider value={C}>
<div style={{minHeight:"100vh",background:C.bg,color:C.wht,fontFamily:"'Segoe UI',system-ui,sans-serif",display:"flex",flexDirection:"column"}}>
<Toast toasts={toasts}/>
<div style={{background:C.card,borderBottom:"1px solid "+C.bdr,padding:"0 12px",display:"flex",alignItems:"center",justifyContent:"space-between",height:54,position:"sticky",top:0,zIndex:100,flexShrink:0,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.05)":"none"}}>
<div style={{display:"flex",alignItems:"center",gap:10}}>
<button onClick={()=>setSideOpen(!sideOpen)} style={{background:"none",border:"none",color:C.gl2,fontSize:22,cursor:"pointer",padding:"8px",borderRadius:6}}>☰</button>
<div style={{display:"flex",alignItems:"center",gap:8}}>
{data.company?.logo?<img src={data.company.logo} style={{height:34,objectFit:"contain"}} alt="logo"/>:<LPGLogo size={28}/>}
{!mobile&&<div><div style={{fontSize:14,fontWeight:900,color:C.blt,lineHeight:1}}>PT. HOE TRANG SA</div><div style={{fontSize:10,color:C.gl2}}>Distributor LPG Pertamina</div></div>}
</div>
</div>
<div style={{display:"flex",alignItems:"center",gap:8}}>
{bonAlerts>0&&<div style={{background:C.mode==="dark"?"#3D1A05":"#FFEDD5",border:"1px solid "+C.olt,borderRadius:20,padding:"3px 9px",fontSize:11,color:C.olt,fontWeight:700}}>⚠️ {bonAlerts}</div>}
<button onClick={themeToggle} style={{background:C.nav,border:"1px solid "+C.bdr,color:C.gl2,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:14}}>{theme==="light"?"🌙":"☀️"}</button>
<button onClick={handlePush} title="Simpan ke Google Sheets" style={{background:syncStatus==="pushing"?C.org:syncStatus==="ok"?C.grn:syncStatus==="error"?C.rdk:C.nav,border:"1px solid "+C.bdr,color:syncStatus==="idle"?C.gl2:"#fff",borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13,fontWeight:700}}>{syncStatus==="pushing"?"⏳":syncStatus==="pulling"?"⏳":"☁️"}</button>
<button onClick={handlePull} title="Ambil data dari Google Sheets" style={{background:C.nav,border:"1px solid "+C.bdr,color:C.gl2,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13}}>⬇️</button>
<div style={{textAlign:"right"}}><div style={{fontSize:12,fontWeight:700,color:C.wht}}>{user.nama.split(" ")[0]}</div>{!mobile&&<div style={{fontSize:10,color:C.gl2}}>{user.posisi}</div>}</div>
<button onClick={()=>setUser(null)} style={{background:C.rdk,border:"none",borderRadius:8,color:"#FFF",fontSize:11,padding:"8px 10px",cursor:"pointer",fontWeight:700}}>⏻</button>
</div>
</div>
<div style={{flex:1,position:"relative",minHeight:0}}>
{sideOpen&&<div onClick={()=>setSideOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:90}}/>}
<div style={{width:220,background:C.card,borderRight:"1px solid "+C.bdr,position:"fixed",top:54,bottom:0,left:0,transform:sideOpen?"translateX(0)":"translateX(-100%)",transition:"transform .22s",zIndex:95,overflowY:"auto"}}>
<div style={{padding:"10px 14px 4px",fontSize:10,color:C.gry,fontWeight:700,letterSpacing:1}}>MENU</div>
{tabs.map(t=><button key={t.id} onClick={()=>{setTab(t.id);setSideOpen(false);}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"12px 16px",background:curTab?.id===t.id?(C.mode==="dark"?"rgba(41,128,185,.2)":"#DBEAFE"):"none",border:"none",color:curTab?.id===t.id?C.blt:C.gl2,fontSize:13,fontWeight:curTab?.id===t.id?700:400,cursor:"pointer",textAlign:"left",borderLeft:curTab?.id===t.id?"3px solid "+C.blt:"3px solid transparent"}}><span style={{fontSize:15}}>{t.icon}</span>{t.label}</button>)}
<div style={{padding:"10px 14px",marginTop:8,borderTop:"1px solid "+C.bdr}}><div style={{fontSize:11,color:C.gl2}}>Login sebagai</div><div style={{fontSize:13,fontWeight:700,color:C.wht}}>{user.nama}</div><div style={{fontSize:10,color:C.gl2}}>{user.posisi}</div></div>
</div>
<div style={{overflowY:"auto",height:"100%",padding:mobile?"10px":"14px",boxSizing:"border-box"}}>
<div style={{display:"flex",gap:5,marginBottom:14,overflowX:"auto",paddingBottom:4}}>
{tabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{background:curTab?.id===t.id?C.blu:C.card,color:curTab?.id===t.id?"#FFF":C.gl2,border:"1px solid "+(curTab?.id===t.id?C.blt:C.bdr),borderRadius:8,padding:mobile?"8px 10px":"6px 13px",fontSize:mobile?11:12,fontWeight:curTab?.id===t.id?700:400,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>{t.icon}{!mobile&&" "+t.label}</button>)}
</div>
<div style={{maxWidth:980,margin:"0 auto"}}>{renderContent()}</div>
</div>
</div>
{inv&&<InvoiceView inv={inv} company={data.company} onClose={()=>setInv(null)}/>}
</div>
</ThemeCtx.Provider>;
}
// === SELESAI BAGIAN 4 — FILE v4 LENGKAP ===
