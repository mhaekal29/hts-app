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
(function(){
if(typeof document!=="undefined"){
if(!document.getElementById("pjs-font")){var l=document.createElement("link");l.id="pjs-font";l.rel="stylesheet";l.href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap";document.head.appendChild(l);}
if(!document.getElementById("h2c-script")){var s=document.createElement("script");s.id="h2c-script";s.src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";document.head.appendChild(s);}
}})();
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
    tasks.push(gasWrite("stok",[{key:"stok",val:JSON.stringify({stock:data.stock,stokKosong:data.stokKosong,totalTabung:data.totalTabung,stokHarian:data.stokHarian,stockLog:data.stockLog,modalHistory:data.modalHistory,hetPrices:data.hetPrices,counters:data.counters,theme:data.theme})}]));
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
stock:{"5.5 kg":0,"12 kg":0,"50 kg":0},stokKosong:{"5.5 kg":0,"12 kg":0,"50 kg":0},totalTabung:{"5.5 kg":0,"12 kg":0,"50 kg":0},
stockLog:[],doList:[],modalHistory:[],hetPrices:{},titipList:[],
penjualan:[],bon:[],pengeluaran:[],employees:DEF_EMP.slice(),
tutupBuku:[],pelanggan:[],setoranSales:[],setoranLog:[],setoranBank:[],kas:{},saldoAwalBank:{BSI:{nominal:0,tanggal:""},BCA:{nominal:0,tanggal:""}},absensi:[],ambilan:[],payrollLog:[],
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

// ── buildStokHarian: rekonstruksi tabel stok per hari dalam sebulan ──
function buildStokHarian(data,bulan){
  var dim=daysInMonth(bulan);
  var rows=[];
  // Stok inject manual (titik awal)
  var injectMap={};
  (data.stokHarian||[]).forEach(function(r){injectMap[r.tanggal]=r;});

  // State berjalan per ukuran
  var curIsi={};var curTK={};
  SIZES.forEach(function(s){curIsi[s]=null;curTK[s]=null;});

  for(var d=1;d<=dim;d++){
    var tgl=bulan+"-"+String(d).padStart(2,"0");
    var dayName=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][new Date(tgl+"T00:00:00").getDay()];

    // Cek inject manual untuk hari ini
    var inject=injectMap[tgl];
    if(inject){
      SIZES.forEach(function(s){curIsi[s]=Number(inject["isi_"+s]||0);curTK[s]=Number(inject["tk_"+s]||0);});
    }

    // Stok awal = nilai berjalan (null kalau belum ada inject)
    var awalIsi={};var awalTK={};
    SIZES.forEach(function(s){awalIsi[s]=curIsi[s]!==null?curIsi[s]:0;awalTK[s]=curTK[s]!==null?curTK[s]:0;});

    // Tabung Masuk: DO diterima + Return/Pancung hari ini
    var doHari=(data.doList||[]).filter(function(d_){return d_.tanggal===tgl&&(d_.status||"diterima")==="diterima";});
    // Mutasi manual masuk (Return, Pancung, Beli Tabung)
    var mutasiMasuk=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Masuk")||l.tanggal===tgl&&(l.jenis||"").includes("Return")||l.tanggal===tgl&&(l.jenis||"").includes("Pancung");});

    var masukIsi={};var masukTK={};
    SIZES.forEach(function(s){
      var doQty=doHari.filter(function(d_){return d_.ukuran===s;}).reduce(function(a,d_){return a+Number(d_.qty||0);},0);
      // Return/Pancung: +isi dari kosong (kosong berkurang seperti DO)
      var retPancQty=mutasiMasuk.filter(function(l){return l.ukuran===s&&((l.jenis||"").includes("Return")||(l.jenis||"").includes("Pancung"));}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // Mutasi masuk lain (tidak termasuk Return/Pancung)
      var mutQtyLain=mutasiMasuk.filter(function(l){return l.ukuran===s&&!(l.jenis||"").includes("Return")&&!(l.jenis||"").includes("Pancung");}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      masukIsi[s]=doQty+retPancQty+mutQtyLain;
      masukTK[s]=doQty+retPancQty;// DO dan Return/Pancung sama-sama kurangi kosong
    });

    // Tabung Keluar: Penjualan + Mutasi Manual rusak/hilang
    var penjHari=(data.penjualan||[]).filter(function(p){return p.tanggal===tgl;});
    var mutasiKeluar=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&((l.jenis||"").includes("Rusak")||(l.jenis||"").includes("Isi Hilang")||(l.jenis||"").includes("Tbg+Isi Hilang"));});
    // Tabung Kosong Hilang: -kosong -total
    var mutasiTbgKosongHilang=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Tbg Kosong Hilang");});
    // Tbg+Isi Hilang: -isi -kosong -total
    var mutasiTbgIsiHilang=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Tbg+Isi Hilang");});
    // Beli Tabung dari Konsumen: +kosong, +total
    var mutasiBeli=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Beli Tabung");});

    var keluarIsi={};var keluarTK={};
    SIZES.forEach(function(s){
      var refillQty=0;var tbgIsiQty=0;
      penjHari.forEach(function(p){
        (p.items||[]).forEach(function(it){
          if(it.ukuran!==s)return;
          var q=Number(it.qty||0);
          if(it.jenis==="Tabung+Isi"){tbgIsiQty+=q;}
          else{refillQty+=q;}
        });
      });
      var rusakQty=mutasiKeluar.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      var beliQty=mutasiBeli.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // keluarIsi: Refill + TbgIsi + Rusak (-isi)
      var tbgKosongHilangQty=mutasiTbgKosongHilang.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      var tbgIsiHilangQty=mutasiTbgIsiHilang.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // keluarIsi tambah tbgIsiHilang
      keluarIsi[s]=refillQty+tbgIsiQty+rusakQty+tbgIsiHilangQty;
      // keluarTK: tbgIsi(-kosong), refill(+kosong), rusak(+kosong), beli(+kosong), tbgKosongHilang(-kosong), tbgIsiHilang(-kosong)
      keluarTK[s]=tbgIsiQty-refillQty-rusakQty-beliQty+tbgKosongHilangQty+tbgIsiHilangQty;
    });

    // Stok Akhir
    var akhirIsi={};var akhirTK={};
    SIZES.forEach(function(s){
      akhirIsi[s]=Math.max(0,awalIsi[s]+masukIsi[s]-keluarIsi[s]);
      akhirTK[s]=Math.max(0,awalTK[s]-masukTK[s]+(-keluarTK[s]));
    });

    // Titip per hari: transaksi titip/tarik yang tanggalnya <= tgl ini
    var titipSnap={};
    SIZES.forEach(function(s){
      var bal=0;
      (data.titipList||[]).filter(function(t){return(t.tanggal||"")<=tgl;}).forEach(function(t){
        var items=t.items&&t.items.length>0?t.items:(t.ukuran===s?[{ukuran:s,qty:t.qty}]:[]);
        var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;
        items.filter(function(i){return i.ukuran===s;}).forEach(function(i){bal+=m*Number(i.qty||0);});
      });
      titipSnap[s]=Math.max(0,bal);
    });

    // Total per ukuran
    var total={};
    SIZES.forEach(function(s){total[s]=akhirIsi[s]+akhirTK[s]+titipSnap[s];});

    // Cek ada transaksi hari ini
    var adaTransaksi=doHari.length>0||penjHari.length>0;

    rows.push({tgl,dayName,d,adaTransaksi,inject:!!inject,
      awalIsi,awalTK,masukIsi,masukTK,keluarIsi,keluarTK,
      akhirIsi,akhirTK,titipSnap,total});

    // Update state berjalan
    SIZES.forEach(function(s){curIsi[s]=akhirIsi[s];curTK[s]=akhirTK[s];});
  }
  return rows;
}
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

// Buat nama file otomatis
function makeFileName(type,label1,label2,ext){
  var d=new Date();
  var tgl=d.getDate().toString().padStart(2,"0")+"-"+["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Ags","Sep","Okt","Nov","Des"][d.getMonth()]+"-"+d.getFullYear();
  var clean=function(s){return(s||"").toUpperCase().replace(/[^A-Z0-9]/g,"-").replace(/-+/g,"-").replace(/^-|-$/g,"").slice(0,30);};
  if(type==="inv")return clean(label1)+"_"+clean(label2)+"_"+tgl+"."+ext;
  if(type==="slip")return clean(label1)+"_Slip-Gaji_"+clean(label2)+"."+ext;
  if(type==="tb")return "Tutup-Buku_"+clean(label1)+"."+ext;
  if(type==="do")return "Laporan-DO_"+clean(label1)+"."+ext;
  return clean(label1)+"_"+tgl+"."+ext;
}

// Cetak PDF dengan saran nama file
function doPrint(id){
var el=document.getElementById(id);
if(!el){alert("Elemen tidak ditemukan: "+id);return;}
var html=el.innerHTML;
var pw=window.open("","_blank","width=900,height=700");
if(!pw){alert("Popup diblokir browser. Izinkan popup untuk mencetak.");return;}
pw.document.write(
"<!DOCTYPE html><html><head><meta charset='utf-8'>"+
"<title>Cetak Laporan</title>"+
"<style>"+
"*{box-sizing:border-box;margin:0;padding:0;}"+
"body{font-family:'Plus Jakarta Sans',Arial,sans-serif;color:#111;background:white;padding:12mm;}"+
"table{width:100%;border-collapse:collapse;margin-bottom:12px;}"+
"th{background:#0a1f44;color:white;padding:6px 8px;font-size:10px;font-weight:700;text-align:center;border:1px solid #ccc;}"+
"td{padding:5px 8px;font-size:11px;border:1px solid #ddd;vertical-align:top;}"+
"tr{page-break-inside:avoid;}"+
"thead{display:table-header-group;}"+
"@page{margin:12mm;size:A4;}"+
"@media print{body{padding:0;}}"+
"</style></head><body>"+
html+
"</body></html>"
);
pw.document.close();
pw.focus();
setTimeout(function(){pw.print();},600);
}

// Download PNG
function doDownloadPNG(id,fileName,onDone){
var el=document.getElementById(id);
if(!el){alert("Element tidak ditemukan");return;}
if(typeof html2canvas==="undefined"){alert("html2canvas belum dimuat, coba lagi sebentar");return;}
html2canvas(el,{scale:2,useCORS:true,backgroundColor:"#ffffff",logging:false}).then(function(canvas){
  var a=document.createElement("a");
  a.href=canvas.toDataURL("image/png");
  a.download=fileName||"dokumen.png";
  a.click();
  if(onDone)onDone();
}).catch(function(e){alert("Gagal buat PNG: "+e.message);});
}

// Copy PNG ke clipboard
function doCopyPNG(id,onDone){
var el=document.getElementById(id);
if(!el){alert("Element tidak ditemukan");return;}
if(typeof html2canvas==="undefined"){alert("html2canvas belum dimuat, coba lagi sebentar");return;}
html2canvas(el,{scale:2,useCORS:true,backgroundColor:"#ffffff",logging:false}).then(function(canvas){
  canvas.toBlob(function(blob){
    if(!blob){alert("Gagal membuat gambar");return;}
    try{
      navigator.clipboard.write([new ClipboardItem({"image/png":blob})]).then(function(){
        if(onDone)onDone("copy");
        alert("✅ Gambar ter-copy! Sekarang paste ke WhatsApp atau chat.");
      }).catch(function(){
        // Fallback: download saja
        var a=document.createElement("a");
        a.href=canvas.toDataURL("image/png");
        a.download="dokumen.png";
        a.click();
        alert("Copy tidak didukung browser ini, otomatis download.");
      });
    }catch(e){alert("Copy gagal: "+e.message);}
  },"image/png");
}).catch(function(e){alert("Gagal buat PNG: "+e.message);});
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
var isLunas=inv.bonLunas||(inv.metodeBayar!=="BON"&&(inv.metodeBayar||"").toLowerCase().indexOf("bon")<0&&!inv.isBon);
var isSplitInv=!!(inv.splitDetail);
var sdInv=inv.splitDetail||{};
var isGabunganInv=!!(inv.isGabungan);
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
// Untuk invoice gabungan, pakai tglDO per item; biasa pakai tanggal invoice
var tglTampil=isGabunganInv&&it.tglDO?it.tglDO:inv.tanggal;
var tglHari=new Date(tglTampil+"T00:00:00");
var HARI_SHORT=["Min","Sen","Sel","Rab","Kam","Jum","Sab"];
return <tr key={i} style={{background:i%2===0?WHITE:G100}}>
<td style={{padding:"10px 14px",color:G400,fontSize:10,fontWeight:600,lineHeight:1.5,borderBottom:"1px solid "+G200,whiteSpace:"nowrap"}}>{HARI_SHORT[tglHari.getDay()]}<br/>{fDs(tglTampil)}</td>
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
{/* Split payment detail */}
{isSplitInv&&<div style={{marginTop:8,borderTop:"1px dashed #00acc1",paddingTop:6}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,color:G600,marginBottom:4,textTransform:"uppercase"}}>Rincian Pembayaran</div>
{[["💵 Cash",sdInv.cash||0,"#065F46"],["🏦 Transfer "+(inv.splitBank||""),sdInv.tf||0,"#1e3a8a"],["📃 BON (Piutang)",sdInv.bon||0,"#991B1B"]].filter(x=>Number(x[1])>0).map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:600,marginBottom:2}}>
<span style={{color:G600}}>{x[0]}</span><span style={{color:x[2]}}>Rp {Number(x[1]).toLocaleString("id-ID")}</span>
</div>)}
</div>}
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
<div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
<div style={{display:"flex",gap:6,flexWrap:"wrap",justifyContent:"center"}}>
{(()=>{var fn=makeFileName("inv",inv.konsumen,inv.noInv,"");return <>
<button onClick={()=>doPrint("_inv",fn+".pdf")} style={{background:NAVY,color:WHITE,border:"none",padding:"9px 20px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_inv",fn+".png")} style={{background:"#1D6A96",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_inv")} style={{background:"#145A32",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>📋 Copy PNG</button>
<button onClick={onClose} style={{background:"#566573",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>✕ Tutup</button>
</>;})()}
</div>
<div style={{fontSize:10,color:"#888",fontStyle:"italic"}}>💡 Nama file PDF: <b style={{color:"#aaa"}}>{makeFileName("inv",inv.konsumen,inv.noInv,"pdf")}</b></div>
</div>
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
<div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
<div style={{display:"flex",gap:6,flexWrap:"wrap",justifyContent:"center"}}>
{(()=>{var fn=makeFileName("slip",slip.nama,"Gaji-"+(slip.bulan||""),"");return <>
<button onClick={()=>doPrint("_slip",fn+".pdf")} style={{background:SNAVY,color:SWHITE,border:"none",padding:"9px 18px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_slip",fn+".png")} style={{background:"#1D6A96",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_slip")} style={{background:"#145A32",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>📋 Copy PNG</button>
<button onClick={onClose} style={{background:"#566573",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>✕ Tutup</button>
</>;})()}
</div>
<div style={{fontSize:10,color:"#888",fontStyle:"italic"}}>💡 Nama file PDF: <b style={{color:"#aaa"}}>{makeFileName("slip",slip.nama,"Gaji-"+(slip.bulan||""),"pdf")}</b></div>
</div>
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
{(()=>{
var rowsLast=buildStokHarian(data,toMonth()).filter(r=>r.tgl<=td);
var lastRow=rowsLast.length>0?rowsLast[rowsLast.length-1]:null;
return <>{SIZES.map(s=>{
var isi=lastRow?lastRow.akhirIsi[s]:((data.stock||{})[s]||0);
var kosong=lastRow?lastRow.akhirTK[s]:getKosong(data,s);
var titip=lastRow?lastRow.titipSnap[s]:getTitipTotal(data.titipList,s);
var totalS=isi+kosong+titip;
return <Card key={s} style={{marginBottom:0}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12}}>📦 LPG {s}</div>
<div style={{fontSize:14,fontWeight:900,color:C.olt}}>{totalS} tab</div>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:4}}>
{[["Tbg+Isi",isi,C.glt],["Titip",titip,C.blt],["Kosong",kosong,C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"5px 4px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:8,color:C.gl2}}>{x[0]}</div><div style={{fontSize:15,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}
</div>
</Card>;
})}</>;
})()}
</div>
<Card style={{marginBottom:14}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Laporan Stok 7 Hari Terakhir</div>
</div>
{(()=>{
  var bulanIni=toMonth();
  var rows7=buildStokHarian(data,bulanIni).filter(r=>r.tgl<=td).slice(-7);
  var uk=["5.5 kg","12 kg","50 kg"];var ukL=["5,5kg","12kg","50kg"];
  var C2=C;
  return <div style={{overflowX:"auto"}}>
  <table style={{borderCollapse:"collapse",width:"100%",fontSize:10}}>
  <thead><tr style={{background:C2.nav}}>
    <th style={{padding:"5px 8px",color:C2.gl2,textAlign:"left",borderBottom:"1px solid "+C2.bdr,whiteSpace:"nowrap"}}>Hari/Tgl</th>
    {uk.map((s,i)=><th key={s} colSpan={4} style={{padding:"5px 6px",color:"white",textAlign:"center",borderBottom:"1px solid "+C2.bdr,borderLeft:"2px solid "+C2.bdr,background:"#1E3A5F",fontSize:9}}>{ukL[i]}: isi | TK | Titip | Total</th>)}
  </tr></thead>
  <tbody>
  {rows7.map((r,i)=><tr key={r.tgl} style={{background:i%2===0?C2.nav:C2.bg,borderBottom:"1px solid "+C2.bdr}}>
    <td style={{padding:"4px 8px",whiteSpace:"nowrap"}}>
      <div style={{fontWeight:700,color:r.tgl===td?C2.blt:C2.wht,fontSize:11}}>{r.dayName}</div>
      <div style={{fontSize:9,color:C2.gl2}}>{fDs(r.tgl)}</div>
    </td>
    {uk.map(s=>[
      <td key={"i"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.glt,fontWeight:700,borderLeft:"2px solid "+C2.bdr}}>{r.akhirIsi[s]}</td>,
      <td key={"k"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.gl2}}>{r.akhirTK[s]}</td>,
      <td key={"t"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.blt}}>{r.titipSnap[s]}</td>,
      <td key={"o"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.olt,fontWeight:700}}>{r.total[s]}</td>
    ])}
  </tr>)}
  </tbody>
  </table>
  </div>;
})()}
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:10}}>⚡ Akses Cepat</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{[["penjualan","🧾","Penjualan"],["stok","📦","Stok"],["piutang","💳","Piutang"],["absensi","📅","Absensi"],["payroll","💼","Payroll"],["laporan","📊","Laporan"]].map(x=><Btn key={x[0]} onClick={()=>setTab(x[0])} sm>{x[1]} {x[2]}</Btn>)}</div></Card>
</div>;
}

// ─── PENJUALAN v4 (FilterTbl, cetak invoice langsung dari riwayat) ────────────
function PenjualanMod({data,setData,setInv,user,toast}){
var C=useTheme();var mob=useMobile();
var blk={ukuran:"5.5 kg",jenis:"Isi",qty:"",price:""};
var canSelf=PENJUALAN_ROLES.includes(user?.role)&&!["owner","admin","akuntan"].includes(user?.role);
var[f,setF]=useState({tanggal:toDay(),salesId:canSelf?user.id:"",konsumen:"",konsumenId:"",items:[{...blk}],bayar:"cash",bank:"BSI",deadline:"",ket:"",splitDetail:{cash:0,tf:0,bon:0},splitBank:"BSI"});
var[delId,setDelId]=useState(null);
var[editInv,setEditInv]=useState(null);// {entry, form}
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",bayar:""});
var[tglLap,setTglLap]=useState(toDay());
var salesEmp=sortEmp((data.employees||[]).filter(e=>e.aktif&&PENJUALAN_ROLES.includes(e.role)));
var valid=f.items.filter(it=>Number(it.qty)>0&&Number(it.price)>0);
var total=iTotal(valid);var margin=calcMargin(valid,data,f.tanggal);
var kNames=[...new Set([...(data.pelanggan||[]).map(p=>p.nama),...(data.penjualan||[]).map(e=>e.konsumen)].filter(Boolean))];
function onKons(nama){var p=(data.pelanggan||[]).find(x=>x.nama===nama);if(p){setF(pv=>{var newItems=pv.items.map(it=>{var h=(Array.isArray(p.hargaKhusus)?p.hargaKhusus:[]).find(x=>x.ukuran===it.ukuran&&x.jenis===it.jenis);if(h)return{...it,price:String(h.harga)};var het=getHET(data,it.ukuran,it.jenis);return{...it,price:het?String(het):it.price};});return{...pv,konsumen:nama,konsumenId:p.id,items:newItems};});}else setF(pv=>({...pv,konsumen:nama,konsumenId:""}));}
function setProduct(i,ukuran,jenis){setF(p=>{var it=p.items.slice();var newIt={...it[i],ukuran,jenis};var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);if(plg){var h=(plg.hargaKhusus||[]).find(x=>x.ukuran===ukuran&&x.jenis===jenis);if(h){newIt.price=String(h.harga);}else newIt.price=String(getHET(data,ukuran,jenis)||"");}else newIt.price=String(getHET(data,ukuran,jenis)||"");it[i]=newIt;return{...p,items:it};});}
function setItem(i,k,v){setF(p=>{var it=p.items.slice();it[i]={...it[i],[k]:v};return{...p,items:it};});}
function makeInvObj(entry){
var emp=(data.employees||[]).find(e=>e.id===entry.salesId);
var plg=(data.pelanggan||[]).find(x=>x.id===entry.konsumenId);
var sd=entry.splitDetail||{};
var metodeBayar=entry.bayar==="bon"?"BON":entry.bayar==="transfer"?"Transfer "+(entry.bank||""):entry.bayar==="split"?"Split":"Cash";
var splitLabel=entry.bayar==="split"?[Number(sd.cash)>0?"Cash":null,Number(sd.tf)>0?"TF":null,Number(sd.bon)>0?"BON":null].filter(Boolean).join("+"):"";
return{noInv:entry.noInv,tanggal:entry.tanggal,konsumen:entry.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",salesNama:emp?.nama||"",items:(entry.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price)})),total:entry.total,metodeBayar:entry.bayar==="split"?splitLabel:metodeBayar,isBon:entry.bayar==="bon"||(entry.bayar==="split"&&Number(sd.bon)>0&&Number(sd.cash)===0&&Number(sd.tf)===0),splitDetail:entry.splitDetail,splitBank:entry.splitBank||"",catatan:entry.ket||""};
}
function doSave(withPrint){
if(!valid.length||!f.konsumen)return;
var ns={...data.stock};
var nk={...(data.stokKosong||{})};
var na={...(data.totalTabung||{})};
var stokLogs=[];
valid.forEach(it=>{
  var q=Number(it.qty||0);var uk=it.ukuran;
  ns[uk]=Math.max(0,(ns[uk]||0)-q);
  if(it.jenis==="Tabung+Isi"){
    nk[uk]=Math.max(0,(nk[uk]||0)-q);
    na[uk]=Math.max(0,(na[uk]||0)-q);// totalTabung berkurang karena tabung terjual
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Tbg+Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  } else {
    nk[uk]=(nk[uk]||0)+q;
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  }
});
var invInfo=nextInvNo(data,f.tanggal);
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.inv)newCounters.inv={};newCounters.inv[invInfo.key]=invInfo.n;
var isSplit=f.bayar==="split";
var sd=f.splitDetail||{cash:0,tf:0,bon:0};
var entry={id:uid(),noInv:invInfo.no,tanggal:f.tanggal,waktu:new Date().toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"}),salesId:f.salesId,konsumen:f.konsumen,konsumenId:f.konsumenId,items:valid.map(it=>({...it,qty:Number(it.qty),price:Number(it.price)})),total,margin,bayar:f.bayar,bank:isSplit?(Number(sd.tf)>0?f.splitBank:""):f.bank,deadline:f.deadline,ket:f.ket,splitDetail:isSplit?sd:null,splitBank:isSplit?f.splitBank:""};
var nb=(data.bon||[]).slice();
// BON: bayar=bon ATAU split dengan bon portion
if(f.bayar==="bon")nb.unshift({id:uid(),noInv:invInfo.no,tanggal:f.tanggal,konsumen:f.konsumen,konsumenId:f.konsumenId,salesId:f.salesId,items:valid,total,sisaTagihan:total,deadline:f.deadline,status:"belum",pembayaran:[],ket:f.ket,bank:f.bank});
if(isSplit&&Number(sd.bon)>0)nb.unshift({id:uid(),noInv:invInfo.no+"(BON)",tanggal:f.tanggal,konsumen:f.konsumen,konsumenId:f.konsumenId,salesId:f.salesId,items:valid,total:Number(sd.bon),sisaTagihan:Number(sd.bon),deadline:f.deadline,status:"belum",pembayaran:[],ket:"Split payment — BON portion. "+f.ket,bank:""});
setData(d=>({...d,penjualan:[entry,...(d.penjualan||[])],stock:ns,stokKosong:nk,totalTabung:na,bon:nb,counters:newCounters,stockLog:[...stokLogs,...(d.stockLog||[])].slice(0,500)}));
if(withPrint)setInv(makeInvObj(entry));
setF(p=>({...p,konsumen:"",konsumenId:"",items:[{...blk}],ket:"",deadline:"",splitDetail:{cash:0,tf:0,bon:0}}));
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
{key:"bayar",label:"Bayar",render:r=>{var sd=r.splitDetail||{};if(r.bayar==="split"){var lbl=[Number(sd.cash)>0?"Cash":null,Number(sd.tf)>0?"TF":null,Number(sd.bon)>0?"BON":null].filter(Boolean).join("+");return <Bdg color="orange">{lbl}</Bdg>;}return r.bayar==="bon"?<Bdg color="red">BON</Bdg>:r.bayar==="transfer"?<Bdg color="blue">TF</Bdg>:<Bdg color="green">Cash</Bdg>;},filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"},{v:"split",l:"Split"}],width:60},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,width:90,render:r=><div style={{display:"flex",gap:4}}><button onClick={()=>setInv(makeInvObj(r))} title="Cetak Invoice" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button><button onClick={()=>{var ef={...r,items:(r.items||[]).map(it=>({...it,qty:String(it.qty),price:String(it.price)}))};setEditInv({entry:r,form:ef});}} title="Edit Invoice" style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"4px 7px",color:"#FCD34D",cursor:"pointer",fontSize:12}}>✏️</button><button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button></div>},
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
<input type="number" value={it.price} step="1000" onChange={e=>setItem(i,"price",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
{!mob&&<span style={{color:C.glt,fontWeight:700,fontSize:12}}>{it.qty&&it.price?fR(Number(it.qty)*Number(it.price)):"-"}</span>}
<button onClick={()=>setF(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} disabled={f.items.length<=1} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13,padding:"2px 5px",opacity:f.items.length<=1?0.3:1}}>✕</button>
</div>)}
<div style={{padding:"7px 10px",background:C.nav,borderTop:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
<Btn sm color="blue" onClick={()=>setF(p=>({...p,items:[...p.items,{...blk}]}))}>+ Item</Btn>
<span style={{fontSize:13,color:C.gl2}}>Margin: <b style={{color:C.glt}}>{fR(margin)}</b> | Total: <b style={{color:C.wht,fontSize:14}}>{fR(total)}</b></span>
</div>
</div>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu],["bon","📃 BON",C.rdk],["split","✂️ Split",C.olt]].map(x=><button key={x[0]} onClick={()=>setF(p=>({...p,bayar:x[0]}))} style={{background:f.bayar===x[0]?x[2]:C.nav,color:f.bayar===x[0]?"white":C.wht,border:"1px solid "+(f.bayar===x[0]?x[2]:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:13,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{f.bayar==="transfer"&&<div style={{display:"flex",gap:8,marginBottom:10}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,bank:b}))} style={{background:f.bank===b?C.blu:C.nav,color:f.bank===b?"white":C.wht,border:"2px solid "+(f.bank===b?C.blt:C.bdr),borderRadius:8,padding:"6px 16px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
{f.bayar==="bon"&&<Inp label="Deadline" type="date" value={f.deadline} onChange={v=>setF(p=>({...p,deadline:v}))} style={{maxWidth:220}}/>}
{f.bayar==="split"&&(()=>{
var sd=f.splitDetail||{cash:0,tf:0,bon:0};
var totalSplit=(Number(sd.cash)||0)+(Number(sd.tf)||0)+(Number(sd.bon)||0);
var selisih=total-totalSplit;
return <div style={{background:C.nav,border:"2px solid "+C.olt,borderRadius:10,padding:12,marginBottom:10}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:8,fontSize:12}}>✂️ Rincian Split Payment</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
{[["cash","💵 Cash",C.glt],["tf","🏦 Transfer",C.blt],["bon","📃 BON",C.rlt]].map(x=><div key={x[0]}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3,fontWeight:600}}>{x[1]}</div>
<input type="number" value={sd[x[0]]||""} placeholder="0" step="1000"
onChange={e=>setF(p=>({...p,splitDetail:{...(p.splitDetail||{}),  [x[0]]:Number(e.target.value)||0}}))}
style={{background:C.bg,border:"1px solid "+(Number(sd[x[0]])>0?x[2]:C.bdr),borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
</div>)}
</div>
{sd.tf>0&&<div style={{display:"flex",gap:6,marginBottom:8}}>
<span style={{fontSize:11,color:C.gl2,alignSelf:"center"}}>TF ke bank:</span>
{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,splitBank:b}))} style={{background:f.splitBank===b?C.blu:C.nav,color:f.splitBank===b?"white":C.wht,border:"2px solid "+(f.splitBank===b?C.blt:C.bdr),borderRadius:6,padding:"4px 12px",fontWeight:700,cursor:"pointer",fontSize:12}}>{b}</button>)}
</div>}
{sd.bon>0&&<Inp label="Deadline BON" type="date" value={f.deadline} onChange={v=>setF(p=>({...p,deadline:v}))} style={{maxWidth:180,marginBottom:8}}/>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:4}}>
{[["Total Split",fR(totalSplit),C.wht],["Total Invoice",fR(total),C.gl2],["Selisih",fR(selisih),Math.abs(selisih)<1?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.bg,borderRadius:6,padding:"5px 8px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
{Math.abs(selisih)<1&&totalSplit>0&&<div style={{marginTop:6,fontSize:11,color:C.glt,fontWeight:700}}>✅ Split balance!</div>}
{Math.abs(selisih)>=1&&totalSplit>0&&<div style={{marginTop:6,fontSize:11,color:C.rlt}}>⚠️ Selisih Rp {fR(Math.abs(selisih))} — sesuaikan nominal</div>}
</div>;
})()}
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))} placeholder="Catatan opsional"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
<Btn onClick={()=>doSave(false)} color="green" dis={!valid.length||!f.konsumen}>💾 Simpan</Btn>
<Btn onClick={()=>doSave(true)} color="blue" dis={!valid.length||!f.konsumen}>🖨️ Simpan & Cetak Invoice</Btn>
</div>
</Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 Laporan Harian Penjualan</div>
{(()=>{
var penjLap=(data.penjualan||[]).filter(p=>p.tanggal===tglLap);
var totalOmzetLap=penjLap.reduce((a,p)=>a+(p.total||0),0);
var totalMarginLap=penjLap.reduce((a,p)=>a+(p.margin||0),0);
// Kelompok per sales
var sgMap={};
penjLap.forEach(p=>{
var empP=(data.employees||[]).find(e=>e.id===p.salesId);
var sNama=empP?.nama||p.salesNama||p.sales||"GUDANG / KASIR";
if(!sgMap[sNama])sgMap[sNama]={nama:sNama,items:[],omzet:0,margin:0,cash:0,tf:0,bon:0};
sgMap[sNama].items.push(p);
sgMap[sNama].omzet+=(p.total||0);
sgMap[sNama].margin+=(p.margin||0);
var byr=(p.bayar||"").toLowerCase();
if(byr==="split"){var sd2=p.splitDetail||{};sgMap[sNama].cash+=Number(sd2.cash||0);sgMap[sNama].tf+=Number(sd2.tf||0);sgMap[sNama].bon+=Number(sd2.bon||0);}
else if(byr==="cash")sgMap[sNama].cash+=(p.total||0);
else if(byr==="transfer"||byr==="tf")sgMap[sNama].tf+=(p.total||0);
else if(byr==="bon")sgMap[sNama].bon+=(p.total||0);
});
var sgList=Object.values(sgMap);
var totCash=sgList.reduce((a,s)=>a+s.cash,0);
var totTF=sgList.reduce((a,s)=>a+s.tf,0);
var totBon=sgList.reduce((a,s)=>a+s.bon,0);
// Total qty & nominal per ukuran keseluruhan
var totQ55=0,totN55=0,totQ12=0,totN12=0,totQ50=0,totN50=0;
penjLap.forEach(p=>(p.items||[]).forEach(it=>{var q=Number(it.qty||0);var h=Number(it.price||0);if(it.ukuran==="5.5 kg"){totQ55+=q;totN55+=q*h;}else if(it.ukuran==="12 kg"){totQ12+=q;totN12+=q*h;}else if(it.ukuran==="50 kg"){totQ50+=q;totN50+=q*h;}}));
return <>
<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10,flexWrap:"wrap"}}>
<Inp label="Tanggal" type="date" value={tglLap} onChange={setTglLap} style={{maxWidth:170,marginBottom:0}}/>
<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
{[["Total Omzet",fR(totalOmzetLap),"#0a1f44"],["Margin",fR(totalMarginLap),"#15803D"],["Invoice",penjLap.length+" trx","#1D4ED8"]].map(x=><div key={x[0]} style={{background:"#F8FAFC",borderRadius:6,padding:"4px 10px",border:"1px solid #E2E8F0"}}><span style={{fontSize:10,color:"#64748B"}}>{x[0]}: </span><span style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</span></div>)}
</div>
</div>
{penjLap.length===0?<div style={{color:"#94A3B8",fontSize:12,padding:"10px 0",fontStyle:"italic"}}>Tidak ada penjualan pada tanggal ini.</div>:
<div style={{background:"white",borderRadius:8,border:"1px solid #E2E8F0",color:"#111",fontFamily:"Arial,sans-serif",overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:700}}>
<thead>
<tr style={{background:"#0a1f44"}}>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"left",border:"1px solid #1e3a5f",fontSize:10}} rowSpan={2}>Nama Sales</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"left",border:"1px solid #1e3a5f",fontSize:10}} rowSpan={2}>Konsumen</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:10}} colSpan={2}>5,5 kg</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:10}} colSpan={2}>12 kg</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:10}} colSpan={2}>50 kg</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"right",border:"1px solid #1e3a5f",fontSize:10}} rowSpan={2}>Total</th>
<th style={{padding:"5px 7px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:10}} rowSpan={2}>Transaksi</th>
</tr>
<tr style={{background:"#1e3a5f"}}>
{["Qty","Harga","Qty","Harga","Qty","Harga"].map((h,i)=><th key={i} style={{padding:"4px 6px",color:"#93C5FD",fontWeight:700,textAlign:"center",border:"1px solid #2d4a6f",fontSize:9}}>{h}</th>)}
</tr>
</thead>
<tbody>
{sgList.map((sg,gi)=>{
var sgRows=sg.items.map((p,i)=>{
var it55=(p.items||[]).filter(it=>it.ukuran==="5.5 kg");var q55p=it55.reduce((a,it)=>a+Number(it.qty||0),0);var h55p=it55.length>0?Number(it55[0].price||0):0;
var it12=(p.items||[]).filter(it=>it.ukuran==="12 kg");var q12p=it12.reduce((a,it)=>a+Number(it.qty||0),0);var h12p=it12.length>0?Number(it12[0].price||0):0;
var it50=(p.items||[]).filter(it=>it.ukuran==="50 kg");var q50p=it50.reduce((a,it)=>a+Number(it.qty||0),0);var h50p=it50.length>0?Number(it50[0].price||0):0;
var byrl=(p.bayar||"").toLowerCase();
return <tr key={p.id} style={{background:i%2===0?"white":"#f9f9f9",borderBottom:"1px solid #e5e7eb"}}>
{i===0&&<td style={{padding:"4px 7px",fontWeight:700,color:"#0a1f44",border:"1px solid #e5e7eb",verticalAlign:"top",background:"#EFF6FF"}} rowSpan={sg.items.length+1}>{sg.nama}</td>}
<td style={{padding:"4px 7px",color:"#111",border:"1px solid #e5e7eb"}}>{p.konsumen}</td>
<td style={{padding:"4px 7px",textAlign:"center",color:q55p>0?"#15803D":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q55p>0?700:400}}>{q55p||"—"}</td>
<td style={{padding:"4px 7px",textAlign:"right",color:q55p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:10}}>{q55p>0?fR(h55p):"—"}</td>
<td style={{padding:"4px 7px",textAlign:"center",color:q12p>0?"#1D4ED8":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q12p>0?700:400}}>{q12p||"—"}</td>
<td style={{padding:"4px 7px",textAlign:"right",color:q12p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:10}}>{q12p>0?fR(h12p):"—"}</td>
<td style={{padding:"4px 7px",textAlign:"center",color:q50p>0?"#B45309":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q50p>0?700:400}}>{q50p||"—"}</td>
<td style={{padding:"4px 7px",textAlign:"right",color:q50p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:10}}>{q50p>0?fR(h50p):"—"}</td>
<td style={{padding:"4px 7px",textAlign:"right",fontWeight:700,color:"#111",border:"1px solid #e5e7eb"}}>{fR(p.total)}</td>
<td style={{padding:"4px 7px",textAlign:"center",border:"1px solid #e5e7eb",color:byrl==="cash"?"#15803D":byrl==="bon"?"#DC2626":"#1D4ED8",fontWeight:700,fontSize:10}}>{p.bayar}</td>
</tr>;});
// Total laku per sales
var tq55sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var tq12sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var tq50sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var totalLakuRow=<tr key={"tl"+gi} style={{background:"#EFF6FF",fontWeight:700,borderBottom:"2px solid #BFDBFE"}}>
<td style={{padding:"5px 7px",color:"#1D4ED8",fontStyle:"italic",border:"1px solid #BFDBFE",fontSize:10}}>Total Laku</td>
<td style={{padding:"5px 7px",textAlign:"center",color:"#15803D",border:"1px solid #BFDBFE"}}>{tq55sg||"—"}</td>
<td style={{padding:"5px 7px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"5px 7px",textAlign:"center",color:"#1D4ED8",border:"1px solid #BFDBFE"}}>{tq12sg||"—"}</td>
<td style={{padding:"5px 7px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"5px 7px",textAlign:"center",color:"#B45309",border:"1px solid #BFDBFE"}}>{tq50sg||"—"}</td>
<td style={{padding:"5px 7px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"5px 7px",textAlign:"right",color:"#0a1f44",border:"1px solid #BFDBFE"}}>{fR(sg.omzet)}</td>
<td style={{padding:"5px 7px",border:"1px solid #BFDBFE"}}></td>
</tr>;
return [...sgRows, totalLakuRow];
})}
{/* Grand total */}
<tr style={{background:"#0a1f44",fontWeight:700}}>
<td colSpan={2} style={{padding:"6px 8px",color:"white",border:"1px solid #1e3a5f"}}>TOTAL KESELURUHAN</td>
<td style={{padding:"6px 8px",textAlign:"center",color:"#86EFAC",border:"1px solid #1e3a5f"}}>{totQ55||"—"}</td>
<td style={{padding:"6px 8px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"6px 8px",textAlign:"center",color:"#93C5FD",border:"1px solid #1e3a5f"}}>{totQ12||"—"}</td>
<td style={{padding:"6px 8px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"6px 8px",textAlign:"center",color:"#FCD34D",border:"1px solid #1e3a5f"}}>{totQ50||"—"}</td>
<td style={{padding:"6px 8px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"6px 8px",textAlign:"right",color:"white",fontSize:13,border:"1px solid #1e3a5f"}}>{fR(totalOmzetLap)}</td>
<td style={{padding:"6px 8px",border:"1px solid #1e3a5f"}}></td>
</tr>
<tr style={{background:"#0f2744"}}>
<td colSpan={2} style={{padding:"5px 8px",color:"#93C5FD",fontSize:10,border:"1px solid #1e3a5f"}}>Cash: {fR(totCash)} &nbsp;|&nbsp; TF: {fR(totTF)} &nbsp;|&nbsp; BON: {fR(totBon)}</td>
<td colSpan={8} style={{padding:"5px 8px",color:"#6B7280",fontSize:10,border:"1px solid #1e3a5f"}}>* BON = piutang baru, belum masuk kas</td>
</tr>
</tbody>
</table>
</div>}
</>;
})()}

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

{editInv&&(()=>{
var ef=editInv.form;
var setEf=newF=>setEditInv(prev=>({...prev,form:typeof newF==="function"?newF(prev.form):newF}));
var setItem2=(i,k,v)=>setEf(p=>{var its=[...p.items];its[i]={...its[i],[k]:v};return{...p,items:its};});
var valid2=(ef.items||[]).filter(it=>Number(it.qty)>0&&Number(it.price)>0);
var total2=iTotal(valid2);
var salesEmpE=sortEmp((data.employees||[]).filter(e=>e.aktif&&PENJUALAN_ROLES.includes(e.role)));
var kNamesE=[...new Set([...(data.pelanggan||[]).map(p=>p.nama),...(data.penjualan||[]).map(e=>e.konsumen)].filter(Boolean))];
function saveEdit(){
var updatedEntry={...editInv.entry,...ef,items:valid2.map(it=>({...it,qty:Number(it.qty),price:Number(it.price)})),total:total2,margin:calcMargin(valid2,data,ef.tanggal),editLog:[...(editInv.entry.editLog||[]),{by:user?.nama||"",at:new Date().toISOString(),before:{konsumen:editInv.entry.konsumen,bayar:editInv.entry.bayar,total:editInv.entry.total,items:editInv.entry.items},note:"Diedit"}]};
// Reverse stok lama lalu apply stok baru
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalTabung||{})};
// Reverse stok lama
(editInv.entry.items||[]).forEach(it=>{var q=Number(it.qty||0);var s=it.ukuran;if(it.jenis==="Tabung+Isi"){ns[s]=(ns[s]||0)+q;nk[s]=(nk[s]||0)+q;na[s]=(na[s]||0)+q;}else{ns[s]=(ns[s]||0)+q;nk[s]=Math.max(0,(nk[s]||0)-q);}});
// Apply stok baru
valid2.forEach(it=>{var q=Number(it.qty||0);var s=it.ukuran;if(it.jenis==="Tabung+Isi"){ns[s]=Math.max(0,(ns[s]||0)-q);nk[s]=Math.max(0,(nk[s]||0)-q);na[s]=Math.max(0,(na[s]||0)-q);}else{ns[s]=Math.max(0,(ns[s]||0)-q);nk[s]=(nk[s]||0)+q;}});
// Jika bayar bon berubah: update bon record
var newBon=(data.bon||[]).map(b=>{if(b.noInv===editInv.entry.noInv&&b.konsumen===editInv.entry.konsumen){return{...b,konsumen:ef.konsumen,konsumenId:ef.konsumenId||b.konsumenId,total:total2,sisaTagihan:total2,items:valid2};}return b;});
setData(d=>({...d,penjualan:(d.penjualan||[]).map(x=>x.id===editInv.entry.id?updatedEntry:x),stock:ns,stokKosong:nk,totalTabung:na,bon:newBon}));
setEditInv(null);toast("✓ Invoice diperbarui! Stok disesuaikan.");
}
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",zIndex:999,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:680,border:"1px solid "+C.bdr,marginTop:20}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div><div style={{fontWeight:700,color:C.wht,fontSize:14}}>✏️ Edit Invoice — {editInv.entry.noInv}</div>
{(editInv.entry.editLog||[]).length>0&&<div style={{fontSize:10,color:C.gl2,marginTop:2}}>Terakhir diedit: {(editInv.entry.editLog||[]).slice(-1)[0]?.by} · {new Date((editInv.entry.editLog||[]).slice(-1)[0]?.at).toLocaleString("id-ID")}</div>}
</div>
<button onClick={()=>setEditInv(null)} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:20}}>✕</button>
</div>
<div style={{padding:16}}>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
<Inp label="Tanggal" type="date" value={ef.tanggal||""} onChange={v=>setEf(p=>({...p,tanggal:v}))}/>
<Sel label="Sales" value={ef.salesId||""} onChange={v=>setEf(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...salesEmpE.map(e=>({v:e.id,l:e.nama}))]}/>
</div>
<div style={{marginBottom:10}}>
<label style={{fontSize:11,color:C.gl2,display:"block",marginBottom:4}}>Konsumen</label>
<input list="kl2" value={ef.konsumen||""} onChange={e=>{var v=e.target.value;var p=(data.pelanggan||[]).find(x=>x.nama===v);setEf(f=>({...f,konsumen:v,konsumenId:p?.id||f.konsumenId}));}} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 12px",color:C.wht,fontSize:13,outline:"none",width:"100%"}}/>
<datalist id="kl2">{kNamesE.map(n=><option key={n} value={n}/>)}</datalist>
</div>
{/* Items */}
<div style={{marginBottom:10}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Items</div>
{(ef.items||[]).map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 80px 90px 32px",gap:6,marginBottom:6,alignItems:"center"}}>
<Sel label="" value={it.ukuran||"5.5 kg"} onChange={v=>setItem2(i,"ukuran",v)} opts={SIZES.map(s=>({v:s,l:s}))}/>
<Sel label="" value={it.jenis||"Isi"} onChange={v=>setItem2(i,"jenis",v)} opts={[{v:"Isi",l:"Isi (Refill)"},{v:"Tabung+Isi",l:"Tabung+Isi"}]}/>
<input type="number" value={it.qty} placeholder="Qty" onChange={e=>setItem2(i,"qty",e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<input type="number" value={it.price} step="1000" placeholder="Harga" onChange={e=>setItem2(i,"price",e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<button onClick={()=>setEf(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} style={{background:"transparent",border:"none",color:C.rlt,cursor:"pointer",fontSize:16}}>−</button>
</div>)}
<button onClick={()=>setEf(p=>({...p,items:[...p.items,{ukuran:"5.5 kg",jenis:"Isi",qty:"",price:""}]}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 12px",color:C.gl2,cursor:"pointer",fontSize:12}}>+ Item</button>
<div style={{textAlign:"right",marginTop:6,fontSize:13,fontWeight:700,color:C.wht}}>Total: {fR(total2)}</div>
</div>
{/* Metode bayar */}
<div style={{marginBottom:10}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Metode Bayar</div>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu],["bon","📃 BON",C.rdk],["split","✂️ Split",C.olt]].map(x=><button key={x[0]} onClick={()=>setEf(p=>({...p,bayar:x[0]}))} style={{background:ef.bayar===x[0]?x[2]:C.nav,color:ef.bayar===x[0]?"white":C.wht,border:"1px solid "+(ef.bayar===x[0]?x[2]:C.bdr),borderRadius:8,padding:"6px 12px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{ef.bayar==="transfer"&&<div style={{display:"flex",gap:8,marginTop:8}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setEf(p=>({...p,bank:b}))} style={{background:ef.bank===b?C.blu:C.nav,color:ef.bank===b?"white":C.wht,border:"2px solid "+(ef.bank===b?C.blt:C.bdr),borderRadius:8,padding:"5px 14px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
</div>
<Inp label="Keterangan (opsional)" value={ef.ket||""} onChange={v=>setEf(p=>({...p,ket:v}))} placeholder="Catatan..."/>
{/* Log perubahan */}
{(editInv.entry.editLog||[]).length>0&&<div style={{marginBottom:10,background:C.bg,borderRadius:8,padding:10,border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📋 Log Perubahan</div>
{(editInv.entry.editLog||[]).slice().reverse().map((lg,i)=><div key={i} style={{fontSize:10,color:C.gl2,marginBottom:4,paddingBottom:4,borderBottom:"1px solid "+C.bdr}}>
<b style={{color:C.wht}}>{lg.by}</b> · {new Date(lg.at).toLocaleString("id-ID")} · {lg.note}<br/>
<span style={{color:"#9CA3AF"}}>Sebelum: {lg.before?.konsumen} · {lg.before?.bayar} · {fR(lg.before?.total||0)}</span>
</div>)}
</div>}
<div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:10}}>
<button onClick={()=>setEditInv(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 18px",color:C.gl2,cursor:"pointer",fontWeight:700}}>Batal</button>
<button onClick={saveEdit} style={{background:C.glt,border:"none",borderRadius:8,padding:"9px 18px",color:"white",cursor:"pointer",fontWeight:700,fontSize:13}}>💾 Simpan Perubahan</button>
</div>
</div>
</div>
</div>;
})()}

{delId&&<ConfirmDel msg={"Hapus invoice "+delId.noInv+" ("+delId.konsumen+")? Stok akan dikembalikan."} onCancel={()=>setDelId(null)} onConfirm={()=>{
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalTabung||{})};
(delId.items||[]).forEach(it=>{
  var q=Number(it.qty||0);var uk=it.ukuran;
  ns[uk]=(ns[uk]||0)+q;
  if(it.jenis==="Tabung+Isi"){nk[uk]=(nk[uk]||0)+q;na[uk]=(na[uk]||0)+q;}
  else{nk[uk]=Math.max(0,(nk[uk]||0)-q);}
});
var log={id:uid(),tanggal:delId.tanggal,ukuran:"Multiple",jenis:"Reverse Hapus Inv "+delId.noInv,qty:0,ket:"Stok dikembalikan karena invoice dihapus",sumber:"Hapus",user:user?.nama||""};
setData(d=>({...d,penjualan:(d.penjualan||[]).filter(x=>x.id!==delId.id),stock:ns,stokKosong:nk,totalTabung:na,stockLog:[log,...(d.stockLog||[])].slice(0,500)}));
setDelId(null);toast("✓ Invoice dihapus & stok dikembalikan!");
}}/>}
</div>;
}

// === AKHIR BAGIAN 2 ===
// === BAGIAN 3 DARI 3 ===

// ─── STOK ─────────────────────────────────────────────────────────────────────
function StokMod({data,setData,user,toast}){
var C=useTheme();var[tab,setTab]=useState("rekap");var[ba,setBa]=useState(null);
var[stokBln,setStokBln]=useState(toMonth());
var[showInject,setShowInject]=useState(false);
var[injectTgl,setInjectTgl]=useState(toDay());
var[injectF,setInjectF]=useState({});
function RekapTab(){
var td2=toDay();
var rowsRekap=buildStokHarian(data,toMonth()).filter(r=>r.tgl<=td2);
var lastRow=rowsRekap.length>0?rowsRekap[rowsRekap.length-1]:null;
return <div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:12,marginBottom:14}}>
{SIZES.map(s=>{
var isi=lastRow?lastRow.akhirIsi[s]:((data.stock||{})[s]||0);
var kosong=lastRow?lastRow.akhirTK[s]:getKosong(data,s);
var titip=lastRow?lastRow.titipSnap[s]:getTitipTotal(data.titipList,s);
var totalS=isi+kosong+titip;
return <Card key={s} style={{marginBottom:0}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
<div style={{fontWeight:800,color:C.wht,fontSize:14}}>📦 LPG {s}</div>
<div style={{textAlign:"right"}}><div style={{fontSize:9,color:C.gl2}}>Total Tabung</div><div style={{fontSize:18,fontWeight:900,color:C.olt}}>{totalS}</div></div>
</div>
{totalS>0&&<div style={{height:6,borderRadius:3,background:C.bdr,display:"flex",overflow:"hidden",marginBottom:8}}>
<div style={{width:(isi/totalS*100)+"%",background:C.glt}}/><div style={{width:(titip/totalS*100)+"%",background:C.blt}}/><div style={{width:(kosong/totalS*100)+"%",background:C.gl2}}/>
</div>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:5}}>
{[["Tbg+Isi",isi,C.glt],["Titip",titip,C.blt],["Kosong",kosong,C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"5px 4px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:8,color:C.gl2}}>{x[0]}</div><div style={{fontSize:18,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}
</div>
</Card>;
})}
</div>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Laporan Stok Harian</div>
<div style={{display:"flex",gap:8,alignItems:"center"}}>
<Inp label="" type="month" value={stokBln} onChange={setStokBln} style={{marginBottom:0,maxWidth:160}}/>
<Btn sm color="orange" onClick={()=>setShowInject(!showInject)}>{showInject?"✕ Tutup":"★ Inject Stok Awal"}</Btn>
</div>
</div>
{showInject&&<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid #F59E0B",marginBottom:12}}>
<div style={{fontWeight:700,color:"#F59E0B",marginBottom:8,fontSize:12}}>★ Set Stok Awal Manual</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Tentukan titik awal perhitungan. Sistem hitung otomatis hari berikutnya.</div>
<div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end",marginBottom:8}}>
<Inp label="Tanggal" type="date" value={injectTgl} onChange={setInjectTgl} style={{maxWidth:160,marginBottom:0}}/>
{SIZES.map(s=>[
<Inp key={"ii"+s} label={"Isi "+s} type="number" value={injectF["isi_"+s]||""} onChange={v=>setInjectF(p=>({...p,["isi_"+s]:v}))} style={{maxWidth:90,marginBottom:0}}/>,
<Inp key={"ik"+s} label={"TK "+s} type="number" value={injectF["tk_"+s]||""} onChange={v=>setInjectF(p=>({...p,["tk_"+s]:v}))} style={{maxWidth:90,marginBottom:0}}/>
])}
</div>
<div style={{display:"flex",gap:8}}>
<Btn color="orange" onClick={()=>{
var rec={tanggal:injectTgl};
SIZES.forEach(s=>{rec["isi_"+s]=Number(injectF["isi_"+s]||0);rec["tk_"+s]=Number(injectF["tk_"+s]||0);});
var existing=(data.stokHarian||[]).filter(r=>r.tanggal!==injectTgl);
setData(d=>({...d,stokHarian:[rec,...existing]}));
setShowInject(false);setInjectF({});
toast("✓ Stok awal "+fDs(injectTgl)+" disimpan!");
}}>💾 Simpan</Btn>
<Btn sm color="gray" onClick={()=>{setShowInject(false);setInjectF({});}}>Batal</Btn>
</div>
{(data.stokHarian||[]).length>0&&<div style={{marginTop:8}}>
{(data.stokHarian||[]).sort((a,b)=>b.tanggal.localeCompare(a.tanggal)).map(r=><div key={r.tanggal} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 8px",background:C.bg,borderRadius:6,marginBottom:3,fontSize:10}}>
<span style={{color:"#F59E0B",fontWeight:700}}>★ {fDs(r.tanggal)} — {SIZES.map(s=>"isi "+s+":"+r["isi_"+s]).join(", ")}</span>
<button onClick={()=>setData(d=>({...d,stokHarian:(d.stokHarian||[]).filter(x=>x.tanggal!==r.tanggal)}))} style={{background:C.rdk,border:"none",borderRadius:4,padding:"2px 7px",color:"white",cursor:"pointer",fontSize:10}}>✕</button>
</div>)}
</div>}
</div>}
<TabelStokBulanan data={data} bulan={stokBln}/>
</Card>
</div>;
}
function MutasiTab(){
var[f,setF]=useState({ukuran:"12 kg",jenis:"return_isi",qty:"",ket:"",tanggal:toDay()});var[delId,setDelId]=useState(null);
function save(){
if(!f.qty||Number(f.qty)<=0)return;
var qty=Number(f.qty);var s=f.ukuran;
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalTabung||{})};
var jDesc="";
if(f.jenis==="return_isi"){
  ns[s]=(ns[s]||0)+qty;        // +isi
  nk[s]=Math.max(0,(nk[s]||0)-qty); // -kosong
  jDesc="↩️ Return (+Isi, -Kosong)";// total tetap
}else if(f.jenis==="pancung"){
  ns[s]=(ns[s]||0)+qty;        // +isi
  nk[s]=Math.max(0,(nk[s]||0)-qty); // -kosong
  jDesc="✂️ Pancung (+Isi, -Kosong)";// total tetap
}else if(f.jenis==="beli_tbg"){
  nk[s]=(nk[s]||0)+qty;        // +kosong
  na[s]=(na[s]||0)+qty;        // total +qty
  jDesc="🛒 Beli Tabung (+Kosong, +Total)";
}else if(f.jenis==="rusak"){
  ns[s]=Math.max(0,(ns[s]||0)-qty); // -isi
  nk[s]=(nk[s]||0)+qty;        // +kosong
  jDesc="💥 Rusak/Bocor (-Isi, +Kosong)";// total tetap
}else if(f.jenis==="tbg_kosong_hilang"){
  nk[s]=Math.max(0,(nk[s]||0)-qty); // -kosong
  na[s]=Math.max(0,(na[s]||0)-qty); // total -qty
  jDesc="🕳️ Tbg Kosong Hilang (-Kosong, -Total)";
}else if(f.jenis==="isi_hilang"){
  ns[s]=Math.max(0,(ns[s]||0)-qty); // -isi
  nk[s]=(nk[s]||0)+qty;        // +kosong (tabung masih ada)
  jDesc="👻 Isi Hilang (-Isi, +Kosong)";// total tetap
}else if(f.jenis==="tbgisi_hilang"){
  ns[s]=Math.max(0,(ns[s]||0)-qty); // -isi
  nk[s]=Math.max(0,(nk[s]||0)-qty); // -kosong
  na[s]=Math.max(0,(na[s]||0)-qty); // total -qty
  jDesc="💀 Tbg+Isi Hilang (-Isi, -Kosong, -Total)";
}
var log={id:uid(),tanggal:f.tanggal,ukuran:s,jenis:jDesc,qty,ket:f.ket,user:user?.nama||"",sumber:"Manual"};
setData(d=>({...d,stock:ns,stokKosong:nk,totalTabung:na,stockLog:[log,...(d.stockLog||[])].slice(0,500)}));
setF(p=>({...p,qty:"",ket:""}));
toast("✓ Mutasi dicatat! "+jDesc);
}
return <div><Card><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10}}><Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/><Sel label="Ukuran" value={f.ukuran} onChange={v=>setF(p=>({...p,ukuran:v}))} opts={SIZES}/><Sel label="Jenis" value={f.jenis} onChange={v=>setF(p=>({...p,jenis:v}))} opts={[{v:"return_isi",l:"↩️ Return (+Isi)"},{v:"pancung",l:"✂️ Pancung (+Isi)"},{v:"beli_tbg",l:"🛒 Beli Tabung dr Konsumen (+Tbg)"},{v:"rusak",l:"💥 Rusak/Bocor (-Isi)"},{v:"tbg_kosong_hilang",l:"🕳️ Tbg Kosong Hilang (-Kosong)"},{v:"isi_hilang",l:"👻 Isi Hilang (-Isi)"},{v:"tbgisi_hilang",l:"💀 Tbg+Isi Hilang (-Isi,-Kosong)"}]}/><Inp label="Qty" type="number" value={f.qty} onChange={v=>setF(p=>({...p,qty:v}))}/><Inp label="Ket" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/></div><Btn color="green" onClick={save} dis={!f.qty}>💾 Simpan Mutasi</Btn></Card><Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Log Mutasi</div><RTbl headers={["Tgl","Ukuran","Jenis","Qty","Ket","Aksi"]} rows={(data.stockLog||[]).slice(0,50).map(l=>[fDs(l.tanggal),l.ukuran,l.jenis,l.qty,l.ket||"-",<ActBtns onDel={()=>setDelId(l)}/>])}/></Card>{delId&&<ConfirmDel msg="Hapus log?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,stockLog:(d.stockLog||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}</div>;
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

// Rekap saldo per konsumen — diurutkan per sales pengantar
var rekapTitip=[];
aktifK.forEach(function(k){
  var b=balMap[k];
  // Cari sales pengantar dari transaksi terakhir
  var lastTitip=(data.titipList||[]).filter(t=>t.konsumenNama===k&&t.tipe==="titip").slice(-1)[0];
  var salesNama=lastTitip?.salesNama||lastTitip?.salesPengantar||"(Tanpa Sales)";
  var total=SIZES.reduce(function(a,s){return a+(b[s]||0);},0);
  rekapTitip.push({konsumen:k,salesNama,s55:b["5.5 kg"]||0,s12:b["12 kg"]||0,s50:b["50 kg"]||0,total,alamat:b.alamat||""});
});
rekapTitip.sort(function(a,b){return a.salesNama.localeCompare(b.salesNama)||a.konsumen.localeCompare(b.konsumen);});

return <div>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}><SC label="Konsumen Titip" value={aktifK.length} icon="🏪" color={C.blt}/>{SIZES.map(s=><SC key={s} label={"Titip "+s} value={getTitipTotal(data.titipList,s)+" tab"} icon="📦" color={C.blt}/>)}</div>

{/* Rekap Saldo + Cetak */}
<Card style={{marginBottom:12}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Rekap Saldo Titip per Konsumen</div>
<button onClick={()=>{
var el=document.getElementById("_lap_titip");if(el)el.remove();
var printDiv=document.createElement("div");printDiv.id="_lap_titip";
var tot12=rekapTitip.reduce(function(a,r){return a+r.s12;},0);
var tot55=rekapTitip.reduce(function(a,r){return a+r.s55;},0);
var tot50=rekapTitip.reduce(function(a,r){return a+r.s50;},0);
var totAll=tot12+tot55+tot50;
var html='<div style="font-family:Arial,sans-serif;padding:20px;color:#111;max-width:900px;margin:0 auto">'+
'<div style="text-align:center;border-bottom:3px solid #0a1f44;padding-bottom:10px;margin-bottom:14px">'+
'<div style="font-size:18px;font-weight:900;color:#0a1f44">LAPORAN TABUNG TITIP DI KONSUMEN</div>'+
'<div style="font-size:13px;font-weight:700;color:#0a1f44">'+(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()+'</div>'+
'<div style="font-size:11px;color:#555;margin-top:4px">Hanya saldo aktif (>0) | Dicetak: '+new Date().toLocaleString("id-ID")+'</div>'+
'</div>'+
'<table style="width:100%;border-collapse:collapse;font-size:11px">'+
'<thead><tr style="background:#0a1f44"><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc;width:30px">No</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Konsumen</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Sales</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">5,5 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">12 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">50 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc;font-weight:900">Total</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Alamat</th></tr></thead>'+
'<tbody>';
rekapTitip.forEach(function(r,i){
html+='<tr style="background:'+(i%2===0?'white':'#f9f9f9')+'">'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center">'+(i+1)+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;font-weight:600">'+r.konsumen+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd">'+r.salesNama+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s55>0?'700':'400')+';color:'+(r.s55>0?'#15803D':'#aaa')+'">'+(r.s55||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s12>0?'700':'400')+';color:'+(r.s12>0?'#1D4ED8':'#aaa')+'">'+(r.s12||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s50>0?'700':'400')+';color:'+(r.s50>0?'#D97706':'#aaa')+'">'+(r.s50||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:900;color:#0a1f44">'+r.total+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;font-size:10px;color:#666">'+r.alamat+'</td>'+
'</tr>';});
html+='<tr style="background:#0a1f44;color:white;font-weight:900">'+
'<td colspan="3" style="padding:6px 8px;border:1px solid #ccc">TOTAL ('+rekapTitip.length+' konsumen)</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot55+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot12+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot50+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+totAll+'</td>'+
'<td style="border:1px solid #ccc"></td></tr>'+
'</tbody></table>'+
'<div style="margin-top:16px;font-size:10px;color:#555;text-align:right">'+data.company?.nama+' — '+data.company?.telepon+'</div>'+
'</div>';
printDiv.innerHTML=html;document.body.appendChild(printDiv);
doPrint("_lap_titip");
setTimeout(function(){var e=document.getElementById("_lap_titip");if(e)e.remove();},3000);
}} style={{background:"#0a1f44",color:"white",border:"none",padding:"7px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak Laporan</button>
</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.nav}}>
{["No","Konsumen","Sales","5,5 kg","12 kg","50 kg","Total","Alamat"].map(h=><th key={h} style={{padding:"6px 8px",color:C.gl2,fontWeight:700,textAlign:["12 kg","5,5 kg","50 kg","Total"].includes(h)?"center":"left",borderBottom:"2px solid "+C.bdr,fontSize:11}}>{h}</th>)}
</tr></thead>
<tbody>
{rekapTitip.map((r,i)=><tr key={r.konsumen} style={{borderBottom:"1px solid "+C.bdr,background:i%2===0?C.nav:C.bg}}>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:11}}>{i+1}</td>
<td style={{padding:"5px 8px",fontWeight:700,color:C.wht}}>{r.konsumen}</td>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:11}}>{r.salesNama}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s55>0?700:400,color:r.s55>0?C.glt:C.gl2}}>{r.s55||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s12>0?700:400,color:r.s12>0?C.blt:C.gl2}}>{r.s12||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s50>0?700:400,color:r.s50>0?C.olt:C.gl2}}>{r.s50||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:900,color:C.wht,fontSize:13}}>{r.total}</td>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:10}}>{r.alamat||"—"}</td>
</tr>)}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
<td colSpan={3} style={{padding:"6px 8px",fontWeight:700,color:C.wht}}>TOTAL ({rekapTitip.length} konsumen)</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.glt}}>{rekapTitip.reduce((a,r)=>a+r.s55,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.blt}}>{rekapTitip.reduce((a,r)=>a+r.s12,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.olt}}>{rekapTitip.reduce((a,r)=>a+r.s50,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.wht,fontSize:14}}>{rekapTitip.reduce((a,r)=>a+r.total,0)}</td>
<td></td>
</tr>
</tbody>
</table>
</div>
</Card>
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
<div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>{[["rekap","📊 Rekap"],["mutasi","✏️ Mutasi Manual"],["opname","🔍 Opname"],["laporan","📋 Lap. Harian"],["titip","🏪 Titip/Tarik"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="rekap"&&<RekapTab/>}
{tab==="mutasi"&&<MutasiTab/>}
{tab==="opname"&&<OpnameTab/>}
{tab==="laporan"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📋 Laporan Stok Harian Bulanan</div>
<div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:14}}>
<Inp label="Pilih Bulan" type="month" value={stokBln} onChange={setStokBln} style={{maxWidth:180,marginBottom:0}}/>
<Btn sm color="orange" onClick={()=>setShowInject(!showInject)}>{showInject?"✕ Tutup":"★ Inject Stok Awal"}</Btn>
</div>
{showInject&&<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid #F59E0B",marginBottom:14}}>
<div style={{fontWeight:700,color:"#F59E0B",marginBottom:10,fontSize:12}}>★ Set Stok Awal Manual — Titik Awal Perhitungan</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10}}>Input stok pada tanggal tertentu sebagai titik awal. Sistem akan hitung otomatis hari berikutnya.</div>
<div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",marginBottom:10}}>
<Inp label="Tanggal Awal" type="date" value={injectTgl} onChange={setInjectTgl} style={{maxWidth:160,marginBottom:0}}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8,marginBottom:10}}>
{SIZES.map(s=>[
<Inp key={"ii"+s} label={"Isi "+s} type="number" value={injectF["isi_"+s]||""} onChange={v=>setInjectF(p=>({...p,["isi_"+s]:v}))} placeholder="0"/>,
<Inp key={"ik"+s} label={"TK "+s+" (Kosong)"} type="number" value={injectF["tk_"+s]||""} onChange={v=>setInjectF(p=>({...p,["tk_"+s]:v}))} placeholder="0"/>
])}
</div>
<div style={{display:"flex",gap:8}}>
<Btn color="orange" onClick={()=>{
var rec={tanggal:injectTgl};
SIZES.forEach(s=>{rec["isi_"+s]=Number(injectF["isi_"+s]||0);rec["tk_"+s]=Number(injectF["tk_"+s]||0);});
var existing=(data.stokHarian||[]).filter(r=>r.tanggal!==injectTgl);
setData(d=>({...d,stokHarian:[rec,...existing]}));
setShowInject(false);setInjectF({});
toast("✓ Stok awal "+fDs(injectTgl)+" berhasil disimpan! Tabel akan dihitung ulang dari tanggal ini.");
}}>💾 Simpan Stok Awal</Btn>
<Btn sm color="gray" onClick={()=>{setShowInject(false);setInjectF({});}}>Batal</Btn>
</div>
{(data.stokHarian||[]).length>0&&<div style={{marginTop:10}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:6,fontWeight:700}}>Inject yang tersimpan:</div>
{(data.stokHarian||[]).slice().sort((a,b)=>b.tanggal.localeCompare(a.tanggal)).map(r=><div key={r.tanggal} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:C.bg,borderRadius:6,marginBottom:4,fontSize:11}}>
<div><b style={{color:"#F59E0B"}}>★ {fDs(r.tanggal)}</b><span style={{color:C.gl2,marginLeft:8}}>{SIZES.map(s=>"isi "+s+": "+r["isi_"+s]+", TK: "+r["tk_"+s]).join(" | ")}</span></div>
<button onClick={()=>{var existing=(data.stokHarian||[]).filter(x=>x.tanggal!==r.tanggal);setData(d=>({...d,stokHarian:existing}));}} style={{background:C.rdk,border:"none",borderRadius:5,padding:"2px 8px",color:"white",cursor:"pointer",fontSize:11}}>✕</button>
</div>)}
</div>}
</div>}
<TabelStokBulanan data={data} bulan={stokBln}/>
</Card>
</div>}
{tab==="titip"&&(()=>{try{return <TitipTab/>;}catch(e){return <Card><div style={{color:C.rlt,padding:12}}>Error: {e.message}</div></Card>;}})()}
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
{delId&&<ConfirmDel msg={"Hapus DO "+(delId.trip||"")+"? Stok akan dikembalikan otomatis."} onCancel={()=>setDelId(null)} onConfirm={()=>{
var st=delId.status||"diterima";
var qty=Number(delId.qty||0);var uk=delId.ukuran;
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
if(st==="diterima"){
  ns[uk]=Math.max(0,(ns[uk]||0)-qty);
  nk[uk]=(nk[uk]||0)+qty;
}
var logs=st==="diterima"?[{id:uid(),tanggal:delId.tanggal,ukuran:uk,jenis:"Reverse Hapus DO "+(delId.trip||""),qty,ket:"Stok dikembalikan karena DO dihapus",sumber:"Hapus DO",user:user?.nama||""},...(data.stockLog||[])].slice(0,500):(data.stockLog||[]);
setData(d=>({...d,doList:(d.doList||[]).filter(x=>x.id!==delId.id),stock:ns,stokKosong:nk,stockLog:logs}));
setDelId(null);
toast(st==="diterima"?"✓ DO dihapus & stok dikembalikan!":"✓ DO dihapus.");
}}/>}
</div>;
}

// ─── PIUTANG v4 (filter sama dgn Penjualan, cetak invoice BON) ────────────────
function PiutangMod({data,setData,setInv,toast}){
var C=useTheme();
var[openId,setOpenId]=useState(null);var[delId,setDelId]=useState(null);
var[bF,setBF]=useState({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",status:""});
var[showGabung,setShowGabung]=useState(false);
var[editPayBon,setEditPayBon]=useState(null);// {bon, payIdx} untuk edit/cancel pembayaran
var[gabungPilih,setGabungPilih]=useState([]);// array bon id yang dipilih
var[gabungKons,setGabungKons]=useState("");// filter konsumen untuk gabung
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));

function makeBonInvObj(b){var plg=(data.pelanggan||[]).find(x=>x.id===b.konsumenId);return{noInv:b.noInv||"#HTS/INV/-/-",tanggal:b.tanggal,konsumen:b.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",items:(b.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price)})),total:b.total,metodeBayar:b.status==="lunas"?"BON (LUNAS)":"BON",isBon:b.status!=="lunas",catatan:b.ket||"",bonLunas:b.status==="lunas"};}

// ── Invoice Gabungan Object ──
function makeGabungInvObj(bons,noInvBaru){
var plg=(data.pelanggan||[]).find(x=>x.id===bons[0]?.konsumenId);
// items: satu row per BON, dengan tanggal DO
var items=bons.flatMap(b=>(b.items||[]).map(it=>({...it,qty:Number(it.qty),price:Number(it.price),tglDO:b.tanggal})));
var total=bons.reduce((a,b)=>a+(b.sisaTagihan||b.total||0),0);
var allLunas=bons.every(b=>b.status==="lunas");
var totalSisa=bons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||b.total||0)),0);
var totalAll=bons.reduce((a,b)=>a+(b.total||0),0);
return{noInv:noInvBaru,tanggal:toDay(),konsumen:bons[0]?.konsumen||"",kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",items,total:totalAll,sisaTagihan:totalSisa,metodeBayar:allLunas?"BON (LUNAS)":"BON (GABUNGAN)",isBon:!allLunas,isGabungan:true,bonLunas:allLunas,bonAsal:bons.map(b=>b.id),catatan:"Invoice gabungan dari "+bons.length+" BON"};
}

function bayar(b){
if(!bF.nominal)return;var nom=Number(bF.nominal);
var newSisa=Math.max(0,(b.sisaTagihan||0)-nom);var st=newSisa===0?"lunas":"sebagian";
var payRec={id:uid(),tanggal:toDay(),jumlah:nom,metode:bF.metode,bank:bF.metode==="transfer"?bF.bank:"",salesPenerimaId:bF.salesPenerimaId,salesPenerimaNama:salesList.find(e=>e.id===bF.salesPenerimaId)?.nama||""};
// Kalau BON gabungan dan lunas → lunasi semua BON asal sekaligus
var newBon=(data.bon||[]).map(x=>{
  if(x.id===b.id)return{...x,sisaTagihan:newSisa,status:st,pembayaran:[...(x.pembayaran||[]),payRec]};
  // Kalau ini BON asal dari gabungan yang baru lunas
  if(st==="lunas"&&b.isGabungan&&(b.bonAsal||[]).includes(x.id))return{...x,status:"digabung",sisaTagihan:0};
  return x;
});
var newSet=data.setoranSales||[];
if(bF.metode==="cash"&&bF.salesPenerimaId)newSet=[{id:uid(),tanggal:toDay(),salesId:bF.salesPenerimaId,sumber:"piutang",refId:b.id,nominal:nom,disetor:false,konsumen:b.konsumen},...newSet];
setData(d=>({...d,bon:newBon,setoranSales:newSet}));
setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});setOpenId(null);toast("✓ Pembayaran tercatat!");
}

// ── Gabung Invoice ──
function doGabung(){
if(gabungPilih.length<2){toast("Pilih minimal 2 BON untuk digabung");return;}
var bons=(data.bon||[]).filter(b=>gabungPilih.includes(b.id));
var invInfo=nextInvNo(data,toDay());
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.inv)newCounters.inv={};newCounters.inv[invInfo.key]=invInfo.n;
var noInvBaru=invInfo.no;
// Tandai BON asal sebagai "digabung"
var newBon=(data.bon||[]).map(b=>{
  if(!gabungPilih.includes(b.id))return b;
  // BON lunas → tetap lunas (jangan ubah status)
  if(b.status==="lunas")return b;
  // BON belum/sebagian → tandai digabung
  return{...b,status:"digabung",sisaTagihan:b.sisaTagihan||b.total||0};
});
// Buat BON gabungan baru
// Total yang masih perlu dibayar (BON lunas tidak dihitung)
var totalGabung=bons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||b.total||0)),0);
// Total keseluruhan untuk info
var totalGabungAll=bons.reduce((a,b)=>a+(b.total||0),0);
var semuaLunas=bons.every(b=>b.status==="lunas");
var bonGabung={id:uid(),noInv:noInvBaru,tanggal:toDay(),konsumen:bons[0].konsumen,konsumenId:bons[0].konsumenId,salesId:bons[0].salesId,items:bons.flatMap(b=>(b.items||[]).map(it=>({...it,qty:Number(it.qty),price:Number(it.price),tglDO:b.tanggal,statusBon:b.status}))),total:totalGabungAll,sisaTagihan:totalGabung,status:semuaLunas?"lunas":"belum",pembayaran:[],isGabungan:true,bonAsal:gabungPilih,ket:"Gabungan dari: "+bons.map(b=>b.noInv).join(", ")};
newBon=[bonGabung,...newBon];
setData(d=>({...d,bon:newBon,counters:newCounters}));
toast("✓ Invoice gabungan dibuat: "+noInvBaru);
setShowGabung(false);setGabungPilih([]);setGabungKons("");
// Tampilkan invoice gabungan
setInv(makeGabungInvObj(bons,noInvBaru));
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

// BON aktif untuk modal gabung
var bonAktifGabung=(data.bon||[]).filter(b=>b.status!=="digabung").filter(b=>!gabungKons||b.konsumen.toLowerCase().includes(gabungKons.toLowerCase()));

var cols=[
{key:"tanggal",label:"Tgl",render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},
{key:"konsumen",label:"Konsumen",render:r=><b style={{color:C.wht,display:"block",minWidth:120}}>{r.konsumen}</b>,filterable:true,width:127},
{key:"salesNama",label:"Sales",filterable:true},
{key:"deadline",label:"Jatuh Tempo",render:r=>r.deadline?<span style={{color:r.dl<0&&r.status!=="lunas"?C.rlt:r.dl<=3&&r.status!=="lunas"?C.olt:C.gl2}}>{fDs(r.deadline)}{r.status!=="lunas"&&r.dl!=null?" ("+(r.dl<0?Math.abs(r.dl)+"h LEWAT":r.dl+"h)"):""}</span>:"-",filterable:false},
{key:"total",label:"Total",render:r=>fR(r.total),filterable:false},
{key:"sisaTagihan",label:"Sisa",render:r=><b style={{color:r.status==="lunas"?C.glt:r.status==="digabung"?C.gl2:C.rlt}}>{fR(r.sisaTagihan)}</b>,filterable:false},
{key:"status",label:"Status",render:r=>r.status==="lunas"?<Bdg color="green">Lunas</Bdg>:r.status==="sebagian"?<Bdg color="orange">Sebagian</Bdg>:r.status==="digabung"?<Bdg color="gray">Digabung</Bdg>:r.isGabungan?<Bdg color="blue">Gabungan</Bdg>:<Bdg color="red">Belum</Bdg>,filterable:true,filterType:"select",options:[{v:"lunas",l:"Lunas"},{v:"sebagian",l:"Sebagian"},{v:"belum",l:"Belum"},{v:"digabung",l:"Digabung"}]},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,render:r=><div style={{display:"flex",gap:4}}><button onClick={()=>setInv(makeBonInvObj(r))} title="Cetak Invoice BON" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button>{(r.status==="belum"||r.status==="sebagian")&&<button onClick={()=>{setOpenId(r.id);setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:r.salesId||""});}} style={{background:C.grn,border:"none",borderRadius:6,padding:"4px 7px",color:"white",cursor:"pointer",fontSize:12}}>💳</button>}{(r.pembayaran||[]).length>0&&<button onClick={()=>setEditPayBon({bon:r,payIdx:null})} title="Edit/Cancel Pembayaran" style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"4px 7px",color:"#FCD34D",cursor:"pointer",fontSize:12}}>✏️</button>}<button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button></div>},
];
var bonActive=(data.bon||[]).filter(b=>b.status!=="lunas"&&b.status!=="digabung");
var totPiutang=bonActive.reduce((a,b)=>a+b.sisaTagihan,0);
return <div>
<STitle icon="💳" children="Piutang / BON"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
<SC label="Total Bon Aktif" value={bonActive.length} icon="📃" color={C.olt}/>
<SC label="Total Piutang" value={fR(totPiutang)} icon="💰" color={C.rlt}/>
<SC label="Bon Lunas" value={(data.bon||[]).filter(b=>b.status==="lunas").length} icon="✅" color={C.glt}/>
</div>

{/* ── MODAL GABUNG INVOICE ── */}
{showGabung&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:600,maxHeight:"85vh",overflow:"hidden",display:"flex",flexDirection:"column",border:"1px solid "+C.bdr}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{fontWeight:700,color:C.wht,fontSize:14}}>🔗 Gabung Invoice BON</div>
<button onClick={()=>{setShowGabung(false);setGabungPilih([]);setGabungKons("");}} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:18}}>✕</button>
</div>
<div style={{padding:"12px 18px",borderBottom:"1px solid "+C.bdr}}>
<Inp label="Filter Konsumen" value={gabungKons} onChange={setGabungKons} placeholder="Ketik nama konsumen..." style={{marginBottom:0}}/>

</div>
<div style={{flex:1,overflowY:"auto",padding:"12px 18px"}}>
{bonAktifGabung.length===0?<div style={{color:C.gl2,fontStyle:"italic",fontSize:12}}>Tidak ada BON aktif{gabungKons?" untuk konsumen ini":""}</div>:
bonAktifGabung.map(b=>{
var isPilih=gabungPilih.includes(b.id);
return <div key={b.id} onClick={()=>setGabungPilih(prev=>prev.includes(b.id)?prev.filter(x=>x!==b.id):[...prev,b.id])} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:isPilih?C.nav:C.bg,borderRadius:8,marginBottom:6,border:"2px solid "+(isPilih?C.blt:C.bdr),cursor:"pointer"}}>
<input type="checkbox" checked={isPilih} onChange={()=>{}} style={{width:16,height:16,cursor:"pointer"}}/>
<div style={{flex:1}}>
<div style={{fontWeight:700,color:C.wht,fontSize:12}}>{b.konsumen}</div>
<div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
<span style={{fontSize:10,color:C.gl2}}>{b.noInv} · {fDs(b.tanggal)}</span>
{b.status==="lunas"?<span style={{fontSize:9,fontWeight:700,color:"#065F46",background:"#D1FAE5",borderRadius:10,padding:"1px 6px"}}>✓ Lunas</span>:b.status==="sebagian"?<span style={{fontSize:9,fontWeight:700,color:"#92400E",background:"#FEF3C7",borderRadius:10,padding:"1px 6px"}}>Sebagian</span>:<span style={{fontSize:9,fontWeight:700,color:"#991B1B",background:"#FEE2E2",borderRadius:10,padding:"1px 6px"}}>Belum Bayar</span>}
</div>
<div style={{fontSize:10,color:C.gl2}}>{(b.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ")}</div>
</div>
<div style={{textAlign:"right"}}>
<div style={{fontWeight:700,color:C.rlt,fontSize:13}}>{fR(b.sisaTagihan)}</div>
<div style={{fontSize:9,color:C.gl2}}>sisa tagihan</div>
</div>
</div>;
})}
</div>
{gabungPilih.length>0&&<div style={{padding:"10px 18px",borderTop:"1px solid "+C.bdr,background:C.nav}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
<span style={{fontSize:12,color:C.gl2}}>{gabungPilih.length} BON dipilih</span>
<div style={{textAlign:"right"}}>
{(()=>{
var pilihBons=(data.bon||[]).filter(b=>gabungPilih.includes(b.id));
var ttlAll=pilihBons.reduce((a,b)=>a+(b.total||0),0);
var ttlSisa=pilihBons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||0)),0);
var adaLunas=pilihBons.some(b=>b.status==="lunas");
return <><div style={{fontSize:14,fontWeight:900,color:C.wht}}>Total: {fR(ttlAll)}</div>
{adaLunas&&ttlSisa>0&&<div style={{fontSize:11,color:C.rlt}}>Sisa tagihan: {fR(ttlSisa)}</div>}
{adaLunas&&ttlSisa===0&&<div style={{fontSize:11,color:C.glt}}>✓ Semua sudah lunas</div>}
</>;
})()}
</div>
</div>
<Btn color="blue" onClick={doGabung} dis={gabungPilih.length<2}>🔗 Gabung & Buat Invoice Baru ({gabungPilih.length} BON)</Btn>
</div>}
</div>
</div>}

<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Daftar Bon</div>
<Btn sm color="blue" onClick={()=>setShowGabung(true)}>🔗 Gabung Invoice</Btn>
</div>
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
{/* Tombol Cetak Laporan BON */}
<div style={{marginBottom:10,display:"flex",gap:8,alignItems:"center"}}>
<button onClick={()=>{
var el=document.getElementById("_lap_bon");if(el)el.remove();
// Build print content
var salesGroups={};
rows.filter(b=>b.status!=="lunas").forEach(b=>{var sn=b.salesNama||"(Tanpa Sales)";if(!salesGroups[sn])salesGroups[sn]=[];salesGroups[sn].push(b);});
var totalSisa=rows.filter(b=>b.status!=="lunas").reduce((a,b)=>a+(b.sisaTagihan||0),0);
var filterInfo=[(barFilter.from||barFilter.to)?("Periode: "+(barFilter.from?fDs(barFilter.from):"")+(barFilter.to?" s/d "+fDs(barFilter.to):"")):"",barFilter.salesId?("Sales: "+(salesList.find(e=>e.id===barFilter.salesId)?.nama||"")):"",barFilter.konsumen?("Konsumen: "+barFilter.konsumen):"",barFilter.status?("Status: "+barFilter.status):""].filter(Boolean).join(" | ")||"Semua Data";
var printDiv=document.createElement("div");printDiv.id="_lap_bon";
var html='<div style="font-family:Arial,sans-serif;padding:20px;color:#111;max-width:900px;margin:0 auto">'+
'<div style="text-align:center;border-bottom:3px solid #0a1f44;padding-bottom:10px;margin-bottom:14px">'+
'<div style="font-size:18px;font-weight:900;color:#0a1f44">LAPORAN PIUTANG / BON AKTIF</div>'+
'<div style="font-size:13px;font-weight:700;color:#0a1f44">'+(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()+'</div>'+
'<div style="font-size:11px;color:#555;margin-top:4px">'+filterInfo+' | Dicetak: '+new Date().toLocaleString("id-ID")+'</div>'+
'</div>';
Object.entries(salesGroups).forEach(function([sn,bons]){
var subTotal=bons.reduce(function(a,b){return a+(b.sisaTagihan||0);},0);
html+='<div style="margin-bottom:14px">';
html+='<div style="font-weight:700;font-size:12px;color:#0a1f44;background:#EFF6FF;padding:5px 8px;border-radius:4px;margin-bottom:6px">Sales: '+sn+'</div>';
html+='<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="background:#0a1f44">';
['No','Konsumen','No.Invoice','Tanggal','Jatuh Tempo','Sisa Tagihan','Status'].forEach(function(h){html+='<th style="color:white;padding:5px 7px;text-align:'+(h==='Sisa Tagihan'?'right':'left');if(h==='No')html+=';width:30px';html+='">'+h+'</th>';});
html+='</tr></thead><tbody>';
bons.forEach(function(b,i){
var dl=dLeft(b.deadline);
var dlTxt=b.deadline?(fDs(b.deadline)+(b.status!=="lunas"&&dl!=null?" ("+(dl<0?Math.abs(dl)+"h LEWAT":dl+"h")+")":'')):'—';
html+='<tr style="background:'+(i%2===0?'white':'#f9f9f9')+'">';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+(i+1)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;font-weight:600">'+b.konsumen+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+( b.noInv||'-')+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+fDs(b.tanggal)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;color:'+(dl!==null&&dl<0?'#DC2626':dl!==null&&dl<=3?'#D97706':'#111')+'">'+dlTxt+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;text-align:right;font-weight:700;color:#DC2626">'+fR(b.sisaTagihan||0)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+b.status+'</td>';
html+='</tr>';});
html+='<tr style="background:#DBEAFE;font-weight:700"><td colspan="5" style="padding:5px 7px;border:1px solid #ddd">Sub-total '+sn+'</td><td style="padding:5px 7px;border:1px solid #ddd;text-align:right;color:#DC2626">'+fR(subTotal)+'</td><td style="border:1px solid #ddd"></td></tr>';
html+='</tbody></table></div>';});
html+='<div style="background:#0a1f44;color:white;padding:8px 12px;border-radius:6px;display:flex;justify-content:space-between;font-weight:700;font-size:13px"><span>TOTAL PIUTANG ('+rows.filter(b=>b.status!=="lunas").length+' bon aktif)</span><span>'+fR(totalSisa)+'</span></div>';
html+='</div>';
printDiv.innerHTML=html;
document.body.appendChild(printDiv);
doPrint("_lap_bon");
setTimeout(function(){var e=document.getElementById("_lap_bon");if(e)e.remove();},3000);
}} style={{background:"#0a1f44",color:"white",border:"none",padding:"8px 16px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak Laporan BON</button>
<span style={{fontSize:10,color:C.gl2,fontStyle:"italic"}}>Cetak sesuai filter aktif — hanya bon belum lunas</span>
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

{/* ── MODAL EDIT/CANCEL PEMBAYARAN BON ── */}
{editPayBon&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:520,border:"1px solid "+C.bdr,maxHeight:"85vh",overflowY:"auto"}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{fontWeight:700,color:C.wht,fontSize:14}}>✏️ Riwayat Pembayaran — {editPayBon.bon.konsumen}</div>
<button onClick={()=>setEditPayBon(null)} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:20}}>✕</button>
</div>
<div style={{padding:16}}>
<div style={{background:C.nav,borderRadius:8,padding:"8px 12px",marginBottom:12,display:"flex",justifyContent:"space-between"}}>
<span style={{fontSize:12,color:C.gl2}}>Invoice: <b style={{color:C.wht}}>{editPayBon.bon.noInv}</b></span>
<span style={{fontSize:12,color:C.gl2}}>Total: <b style={{color:C.rlt}}>{fR(editPayBon.bon.total)}</b></span>
</div>
{(editPayBon.bon.pembayaran||[]).length===0
?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada pembayaran</div>
:(editPayBon.bon.pembayaran||[]).map((pay,pi)=>{
var isEdit=editPayBon.payIdx===pi;
return <div key={pi} style={{background:isEdit?C.nav:C.bg,borderRadius:8,padding:"10px 12px",marginBottom:8,border:"2px solid "+(isEdit?C.olt:C.bdr)}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:isEdit?8:0}}>
<div>
<div style={{fontSize:13,fontWeight:700,color:C.wht}}>{fR(pay.jumlah||pay.nominal||0)}</div>
<div style={{fontSize:10,color:C.gl2}}>{fDs(pay.tanggal)} · <span style={{color:(pay.metode||"").toLowerCase()==="cash"?C.glt:C.blt,fontWeight:700}}>{pay.metode||"Cash"}</span>{pay.bank?" · "+pay.bank:""}</div>
{pay.salesPenerimaNama&&<div style={{fontSize:10,color:C.gl2}}>Diterima: {pay.salesPenerimaNama}</div>}
</div>
<div style={{display:"flex",gap:5}}>
<button onClick={()=>setEditPayBon(p=>({...p,payIdx:isEdit?null:pi,editForm:{...pay}}))} style={{background:isEdit?C.olt:C.nav,border:"1px solid "+(isEdit?C.olt:C.bdr),borderRadius:6,padding:"4px 9px",color:isEdit?"white":C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>{isEdit?"▲":"✏️ Edit"}</button>
<button onClick={()=>{
var bonId=editPayBon.bon.id;
var bon=(data.bon||[]).find(b=>b.id===bonId)||editPayBon.bon;
var newPays=(bon.pembayaran||[]).filter((_,idx)=>idx!==pi);
var newSisa=Math.max(0,bon.total-newPays.reduce((a,p)=>a+(Number(p.jumlah||p.nominal||0)),0));
var newStatus=newPays.length===0?"belum":newSisa<=0?"lunas":"sebagian";
var log={id:uid(),type:"cancel_pay",by:user?.nama||"Admin",at:new Date().toISOString(),before:{jumlah:pay.jumlah||pay.nominal||0,metode:pay.metode,tanggal:pay.tanggal},note:"Pembayaran dicancel"};
var updBon={...bon,pembayaran:newPays,sisaTagihan:newSisa,status:newStatus,editLog:[...(bon.editLog||[]),log]};
var newBon=(data.bon||[]).map(b=>b.id===bonId?updBon:b);
setData(d=>({...d,bon:newBon}));
setEditPayBon({bon:updBon,payIdx:null,editForm:null});
toast("✓ Pembayaran dicancel. Status: "+newStatus);
}} style={{background:C.rdk,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 9px",color:"white",cursor:"pointer",fontSize:11,fontWeight:700}}>🗑️ Cancel</button>
</div>
</div>
{isEdit&&editPayBon.editForm&&<div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
<div><div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Nominal</div>
<input type="number" value={editPayBon.editForm.jumlah||editPayBon.editForm.nominal||""} onChange={e=>setEditPayBon(p=>({...p,editForm:{...p.editForm,jumlah:Number(e.target.value)}}))} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"7px 9px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/></div>
<div><div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Tanggal</div>
<input type="date" value={editPayBon.editForm.tanggal||""} onChange={e=>setEditPayBon(p=>({...p,editForm:{...p.editForm,tanggal:e.target.value}}))} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"7px 9px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/></div>
</div>
<div style={{marginBottom:8}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:4}}>Metode</div>
<div style={{display:"flex",gap:6}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu]].map(x=><button key={x[0]} onClick={()=>setEditPayBon(p=>({...p,editForm:{...p.editForm,metode:x[0]}}))} style={{background:(editPayBon.editForm.metode||"cash")===x[0]?x[2]:C.nav,color:(editPayBon.editForm.metode||"cash")===x[0]?"white":C.wht,border:"1px solid "+((editPayBon.editForm.metode||"cash")===x[0]?x[2]:C.bdr),borderRadius:7,padding:"5px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{(editPayBon.editForm.metode||"cash")==="transfer"&&<div style={{display:"flex",gap:6,marginTop:6}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setEditPayBon(p=>({...p,editForm:{...p.editForm,bank:b}}))} style={{background:(editPayBon.editForm.bank||"BSI")===b?C.blu:C.nav,color:(editPayBon.editForm.bank||"BSI")===b?"white":C.wht,border:"2px solid "+((editPayBon.editForm.bank||"BSI")===b?C.blt:C.bdr),borderRadius:7,padding:"4px 14px",fontWeight:700,cursor:"pointer",fontSize:12}}>{b}</button>)}</div>}
</div>
<button onClick={()=>{
var bonId=editPayBon.bon.id;
var bon=(data.bon||[]).find(b=>b.id===bonId)||editPayBon.bon;
var pi2=editPayBon.payIdx;
var ef2=editPayBon.editForm;
if(pi2===null||pi2===undefined||!ef2)return;
var log={id:uid(),type:"edit_pay",by:user?.nama||"Admin",at:new Date().toISOString(),before:{...(bon.pembayaran||[])[pi2]},note:"Pembayaran diedit"};
var newPays=(bon.pembayaran||[]).map((p,idx)=>idx===pi2?{...p,...ef2,jumlah:Number(ef2.jumlah||ef2.nominal||p.jumlah||p.nominal||0)}:p);
var newSisa=Math.max(0,bon.total-newPays.reduce((a,p)=>a+(Number(p.jumlah||p.nominal||0)),0));
var newStatus=newSisa<=0?"lunas":newPays.length===0?"belum":"sebagian";
var updBon={...bon,pembayaran:newPays,sisaTagihan:newSisa,status:newStatus,editLog:[...(bon.editLog||[]),log]};
var newBon=(data.bon||[]).map(b=>b.id===bonId?updBon:b);
setData(d=>({...d,bon:newBon}));
setEditPayBon({bon:updBon,payIdx:null,editForm:null});
toast("✓ Pembayaran diperbarui! Status: "+newStatus);
}} style={{background:C.glt,border:"none",borderRadius:8,padding:"7px 18px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12,width:"100%"}}>💾 Simpan Perubahan</button>
</div>}
</div>;})}
{/* Log perubahan */}
{(editPayBon.bon.editLog||[]).length>0&&<div style={{marginTop:10,background:C.bg,borderRadius:8,padding:10,border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📋 Log Perubahan</div>
{(editPayBon.bon.editLog||[]).slice().reverse().map((lg,i)=><div key={i} style={{fontSize:10,color:C.gl2,marginBottom:4,paddingBottom:4,borderBottom:"1px solid "+C.bdr}}>
<b style={{color:C.wht}}>{lg.by}</b> · {new Date(lg.at).toLocaleString("id-ID")} · <span style={{color:lg.type==="cancel_pay"?C.rlt:C.olt}}>{lg.note}</span>
{lg.before&&<><br/><span style={{color:"#9CA3AF"}}>Sebelum: {fR(lg.before.jumlah||0)} · {lg.before.metode} · {fDs(lg.before.tanggal)}</span></>}
</div>)}
</div>}
</div>
</div>
</div>}
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
var[checkedPen,setCheckedPen]=useState({});
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));
var emp=(data.employees||[]).find(e=>e.id===salesId);

// ── DATA HARI INI ──
var penjHari=(data.penjualan||[]).filter(p=>p.salesId===salesId&&p.tanggal===tgl);
var penjCash=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="cash");
var penjTF=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf");
var penjBon=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="bon");
var penjSplit=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="split");
// Hitung nominal per metode — termasuk split
var cashPenjualan=penjCash.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).cash||0),0);
var tfPenjualan=penjTF.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).tf||0),0);
var bonPenjualan=penjBon.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).bon||0),0);

// Bayar BON hari ini
var bonBayarList=(data.bon||[]).flatMap(b=>(b.pembayaran||[]).filter(px=>px.tanggal===tgl&&b.salesId===salesId).map(px=>({...px,konsumen:b.konsumen})));
var bonCash=bonBayarList.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTF=bonBayarList.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var totalBonBayar=bonCash+bonTF;

// Pengeluaran operasional hari ini
var allPenHari=(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var penSales=allPenHari.filter(p=>p.karyawanId===salesId||p.karyawanNama===(emp?.nama||""));
var penLain=allPenHari.filter(p=>p.karyawanId!==salesId&&p.karyawanNama!==(emp?.nama||""));

// Init centang: auto centang atas nama sales, bisa tambah manual
useEffect(()=>{
  if(!salesId)return;
  var cp={};
  allPenHari.forEach(p=>{cp[p.id]=!!(p.karyawanId===salesId||p.karyawanNama===(emp?.nama||""));});
  setCheckedPen(cp);
},[salesId,tgl]);

var totalPenChecked=allPenHari.filter(p=>checkedPen[p.id]).reduce((a,p)=>a+Number(p.nominal||0),0);
var totalCashWajibSetor=Math.max(0,cashPenjualan+bonCash-totalPenChecked);
var totalTunai=DENOMS.reduce((a,d)=>a+Number(pecah[d]||0)*d,0);
var selisih=totalTunai-totalCashWajibSetor;

// qty per ukuran helper
var getQty=(p,uk)=>(p.items||[]).filter(it=>it.ukuran===uk).reduce((a,it)=>a+Number(it.qty||0),0);
var getHarga=(p,uk)=>{var it=(p.items||[]).filter(it=>it.ukuran===uk);return it.length>0?Number(it[0].price||0):0;};

function konfirmasi(){
var newAmb=[...(data.ambilan||[])];
if(selisih<0&&jadikanPinjaman){
  newAmb.unshift({id:uid(),karyawanId:salesId,karyawanNama:emp?.nama||"",nominal:Math.abs(selisih),ket:"Kurang setor "+fDs(tgl),tanggal:tgl});
}
var logEntry={id:uid(),tanggal:tgl,salesId,salesNama:emp?.nama||"",
  cashPenjualan,tfPenjualan,bonPenjualan,bonCash,bonTF,totalBonBayar,
  totalPotongan:totalPenChecked,totalCashWajibSetor,totalTunai,selisih,
  jadikanPinjaman:selisih<0&&jadikanPinjaman};
setData(d=>({...d,ambilan:newAmb,setoranLog:[logEntry,...(d.setoranLog||[])]}));
setPecah(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
setJadikanPinjaman(false);
toast("✓ Setoran dikonfirmasi! Wajib setor: "+fR(totalCashWajibSetor));
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
{/* ── 1. PENJUALAN per konsumen ── */}
<Card style={{border:"1px solid "+C.blt,overflow:"hidden"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:8,fontSize:13}}>📦 Penjualan — {fDs(tgl)}</div>
{penjHari.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada penjualan hari ini</div>:<>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:600}}>
<thead>
<tr style={{background:C.nav}}>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Konsumen</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>5,5 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>12 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>50 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Total</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Transaksi</th>
</tr>
<tr style={{background:C.nav}}>
{["Qty","Harga","Qty","Harga","Qty","Harga"].map((h,i)=><th key={i} style={{padding:"4px 6px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:9}}>{h}</th>)}
</tr>
</thead>
<tbody>
{penjHari.map((p,i)=>{
var q55=getQty(p,"5.5 kg");var h55=getHarga(p,"5.5 kg");
var q12=getQty(p,"12 kg");var h12=getHarga(p,"12 kg");
var q50=getQty(p,"50 kg");var h50=getHarga(p,"50 kg");
var byr=(p.bayar||"").toLowerCase();
return <tr key={p.id} style={{background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{p.konsumen}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q55>0?C.glt:C.gl2,border:"1px solid "+C.bdr}}>{q55||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q55>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q55>0?fR(h55):"—"}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q12>0?C.blt:C.gl2,border:"1px solid "+C.bdr}}>{q12||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q12>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q12>0?fR(h12):"—"}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q50>0?C.olt:C.gl2,border:"1px solid "+C.bdr}}>{q50||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q50>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q50>0?fR(h50):"—"}</td>
<td style={{padding:"4px 8px",textAlign:"right",fontWeight:700,color:C.wht,border:"1px solid "+C.bdr}}>{fR(p.total)}</td>
<td style={{padding:"4px 6px",textAlign:"center",border:"1px solid "+C.bdr,color:byr==="cash"?C.glt:byr==="bon"?C.rlt:C.blt,fontWeight:700,fontSize:10}}>{p.bayar}</td>
</tr>;})}
{/* Total Laku */}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td style={{padding:"5px 8px",color:C.blt,fontStyle:"italic",border:"1px solid "+C.bdr,fontSize:10}}>Total Laku</td>
{(()=>{
var tq55=penjHari.reduce((a,p)=>a+getQty(p,"5.5 kg"),0);
var tq12=penjHari.reduce((a,p)=>a+getQty(p,"12 kg"),0);
var tq50=penjHari.reduce((a,p)=>a+getQty(p,"50 kg"),0);
return <>
<td style={{padding:"5px 6px",textAlign:"center",color:C.glt,border:"1px solid "+C.bdr}}>{tq55||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 6px",textAlign:"center",color:C.blt,border:"1px solid "+C.bdr}}>{tq12||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 6px",textAlign:"center",color:C.olt,border:"1px solid "+C.bdr}}>{tq50||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.wht,border:"1px solid "+C.bdr}}>{fR(penjHari.reduce((a,p)=>a+(p.total||0),0))}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
</>;})()} 
</tr>
</tbody>
</table>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:8}}>
{[["Cash",fR(cashPenjualan),C.glt,penjCash.length+" trx"],["Transfer",fR(tfPenjualan),C.blt,penjTF.length+" trx"],["BON (piutang)",fR(bonPenjualan),"#9CA3AF",penjBon.length+" trx"]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"6px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div><div style={{fontSize:9,color:C.gl2}}>{x[3]}</div></div>)}
</div>
</>}
</Card>

{/* ── 2. BAYAR BON ── */}
<Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:8,fontSize:13}}>💳 Bayar Bon / Piutang</div>
{bonBayarList.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada pembayaran BON hari ini</div>:<>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,marginBottom:8}}>
<thead><tr style={{background:C.nav}}>
{["Konsumen","Nominal","Metode"].map(h=><th key={h} style={{padding:"5px 8px",color:C.gl2,textAlign:h==="Nominal"?"right":"left",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{bonBayarList.map((b,i)=><tr key={i} style={{background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,border:"1px solid "+C.bdr}}>{b.konsumen}</td>
<td style={{padding:"4px 8px",textAlign:"right",fontWeight:700,color:C.wht,border:"1px solid "+C.bdr}}>{fR(b.jumlah||b.nominal||0)}</td>
<td style={{padding:"4px 8px",color:(b.metode||"cash").toLowerCase()==="cash"?C.glt:C.blt,fontWeight:700,border:"1px solid "+C.bdr,fontSize:10}}>{b.metode||"Cash"}</td>
</tr>)}
</tbody>
</table>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
{[["Cash",fR(bonCash),C.glt],["Transfer",fR(bonTF),C.blt],["TOTAL BAYAR BON",fR(totalBonBayar),C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"6px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
</>}
</Card>

{/* ── 3. OPERASIONAL SALES ── */}
<Card style={{border:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:6,fontSize:13}}>💸 Operasional Sales</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>✅ Auto tercentang = atas nama {emp?.nama||"sales ini"}. Klik untuk ubah.</div>
{allPenHari.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada pengeluaran hari ini</div>:
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
{["","Kategori","Keterangan","Atas Nama","Metode","Nominal"].map(h=><th key={h} style={{padding:"5px 7px",color:C.gl2,textAlign:h==="Nominal"?"right":"left",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{allPenHari.map((p,i)=>{
var isMe=p.karyawanId===salesId||p.karyawanNama===(emp?.nama||"");
var checked=!!checkedPen[p.id];
return <tr key={p.id} style={{background:checked?(i%2===0?C.nav:C.bg):(i%2===0?C.bg:"transparent"),borderBottom:"1px solid "+C.bdr,cursor:"pointer",opacity:checked?1:.6}} onClick={()=>setCheckedPen(prev=>({...prev,[p.id]:!prev[p.id]}))}>
<td style={{padding:"4px 7px",border:"1px solid "+C.bdr}}><input type="checkbox" checked={checked} onChange={()=>{}} style={{cursor:"pointer"}}/></td>
<td style={{padding:"4px 7px",color:C.wht,border:"1px solid "+C.bdr,fontWeight:600}}>{p.kategori}</td>
<td style={{padding:"4px 7px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{p.ket||"—"}</td>
<td style={{padding:"4px 7px",border:"1px solid "+C.bdr}}>
<span style={{color:isMe?C.glt:C.gl2,fontSize:10,fontWeight:isMe?700:400}}>{p.karyawanNama||"—"}{isMe?" (auto)":""}</span>
</td>
<td style={{padding:"4px 7px",color:(p.metode||"cash").toLowerCase()==="cash"?C.glt:C.blt,border:"1px solid "+C.bdr,fontWeight:700,fontSize:10}}>{p.metode||"Cash"}</td>
<td style={{padding:"4px 7px",textAlign:"right",fontWeight:700,color:checked?C.rlt:C.gl2,border:"1px solid "+C.bdr}}>{fR(p.nominal)}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td colSpan={5} style={{padding:"5px 7px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:11}}>Total Operasional (yang dicentang)</td>
<td style={{padding:"5px 7px",textAlign:"right",color:C.rlt,border:"1px solid "+C.bdr,fontSize:13}}>{fR(totalPenChecked)}</td>
</tr>
</tbody>
</table>}
</Card>

{/* ── 4. RINGKASAN WAJIB SETOR ── */}
<Card style={{border:"2px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:10,fontSize:13}}>🧾 Ringkasan Wajib Setor Cash</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:12}}>
<tbody>
{[
["Cash Penjualan",cashPenjualan,C.glt,false],
["Bayar BON Cash",bonCash,C.glt,false],
["Pengeluaran Cash (dipotong)","-"+fR(totalPenChecked),C.rlt,false],
["TOTAL CASH WAJIB SETOR",totalCashWajibSetor,C.wht,true],
["Total TF Penjualan (info)",tfPenjualan,"#9CA3AF",false],
["Bayar BON TF (info)",bonTF,"#9CA3AF",false],
["Total BON Piutang Baru (info)",bonPenjualan,"#9CA3AF",false],
].map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:r[3]?C.nav:"transparent"}}>
<td style={{padding:"6px 10px",color:r[3]?C.wht:C.gl2,fontWeight:r[3]?700:400}}>{r[0]}</td>
<td style={{padding:"6px 10px",textAlign:"right",color:r[2],fontWeight:r[3]?800:600,fontSize:r[3]?15:12}}>{typeof r[1]==="string"?r[1]:fR(r[1])}</td>
</tr>)}
</tbody>
</table>
</Card>

{/* ── 5. PECAHAN KAS ── */}
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

{/* ── 6. REKONSILIASI ── */}
<Card style={{border:"2px solid "+(Math.abs(selisih)<1000?C.glt:C.rlt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔄 Rekonsiliasi</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
{[["Wajib Setor",totalCashWajibSetor,C.wht],["Tunai Fisik",totalTunai,C.glt],["Selisih",selisih,selisih>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr,textAlign:"center"}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:Math.abs(selisih)<1000?C.grn:C.rdk,borderRadius:8,marginBottom:12}}>
<span style={{color:"white",fontWeight:700}}>SELISIH {selisih>=0?"LEBIH":"KURANG"}</span>
<b style={{color:"white",fontSize:18}}>{fR(Math.abs(selisih))}</b>
</div>
{selisih<0&&<div style={{marginBottom:8}}>
<label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",padding:"8px 12px",background:C.rdk,borderRadius:8}}>
<input type="checkbox" checked={jadikanPinjaman} onChange={e=>setJadikanPinjaman(e.target.checked)} style={{width:16,height:16}}/>
<span style={{color:"white",fontSize:12,fontWeight:700}}>💳 Jadikan pinjaman karyawan (kurang setor {fR(Math.abs(selisih))})</span>
</label>
</div>}
<Btn color="green" onClick={konfirmasi} dis={!salesId}>✓ Konfirmasi Setoran</Btn>
</Card>
</>}

{riwayat.length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Setoran</div>
<RTbl headers={["Tgl","Sales","Cash Penj","Bayar BON","Potongan","Wajib Setor","Tunai","Selisih"]} rows={riwayat.map(r=>[
fDs(r.tanggal),r.salesNama,
fR(r.cashPenjualan||0),
fR(r.bonCash||r.bonBayarHari||0),
fR(r.totalPotongan||0),
fR(r.totalCashWajibSetor||r.bersihSetor||0),
fR(r.totalTunai||0),
<b style={{color:Math.abs(r.selisih||0)<1000?C.glt:C.rlt}}>{fR(r.selisih||0)}</b>
])}/>
</Card>}
</div>;
}

// ─── TUTUP BUKU v4 (harian + bulanan, PDF/PNG) ────────────────────────────────
// ── TabelStokBulanan: tampil tabel per hari dalam sebulan ──
function TabelStokBulanan({data,bulan,compact}){
  var C=useTheme();
  var rows=buildStokHarian(data,bulan);
  var uk=["12 kg","5.5 kg","50 kg"];
  var ukLabel=["12kg","5,5kg","50kg"];
  var headerGroups=[
    {label:"STOK AWAL",cols:["isi","TK"],color:"#1E40AF"},
    {label:"TABUNG MASUK",cols:["isi","TK"],color:"#065F46"},
    {label:"TABUNG KELUAR",cols:["isi","TK"],color:"#7F1D1D"},
    {label:"STOK AKHIR",cols:["isi","TK","Titip","Total"],color:"#1E3A5F"},
  ];
  var thS=function(c,align){return{padding:"5px 6px",color:"white",fontWeight:700,fontSize:9,textAlign:align||"center",borderRight:"1px solid rgba(255,255,255,.15)",background:c,whiteSpace:"nowrap"};};
  var tdS=function(color,bold){return{padding:"4px 6px",textAlign:"center",color:color||C.wht,fontWeight:bold?700:400,fontSize:11,borderRight:"1px solid "+C.bdr,borderBottom:"1px solid "+C.bdr,whiteSpace:"nowrap"};};

  return <div style={{overflowX:"auto"}}>
  <table style={{borderCollapse:"collapse",fontSize:11,width:"100%",minWidth:900}}>
  <thead>
  <tr>
    <th rowSpan={3} style={{...thS("#0F172A"),padding:"6px 8px",fontSize:10,minWidth:28}}>No</th>
    <th rowSpan={3} style={{...thS("#0F172A"),minWidth:90,textAlign:"left",padding:"6px 8px"}}>Hari/Tanggal</th>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=><th key={g} colSpan={uk.length*2} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),padding:"6px 4px",letterSpacing:.5}}>{g}</th>)}
    <th colSpan={uk.length*4} style={{...thS("#1E3A5F"),padding:"6px 4px",letterSpacing:.5}}>STOK AKHIR</th>
  </tr>
  <tr>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=>uk.map((s,si)=><th key={g+s} colSpan={2} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),fontSize:9}}>{ukLabel[si]}</th>))}
    {uk.map((s,si)=><th key={"akhir"+s} colSpan={4} style={{...thS("#1E3A5F"),fontSize:9}}>{ukLabel[si]}</th>)}
  </tr>
  <tr>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=>uk.map(s=>["isi","TK"].map(c=><th key={g+s+c} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),fontSize:8}}>{c}</th>)))}
    {uk.map(s=>["isi","TK","Titip","Total"].map(c=><th key={"akhir"+s+c} style={{...thS("#1E3A5F"),fontSize:8,fontWeight:c==="Total"?900:700}}>{c}</th>))}
  </tr>
  </thead>
  <tbody>
  {rows.map((r,i)=>{
    var isLibur=!r.adaTransaksi&&!r.inject;
    var bg=i%2===0?C.nav:C.bg;
    return <tr key={r.tgl} style={{background:isLibur?"rgba(100,100,100,.1)":bg}}>
    <td style={tdS(C.gl2)}>{r.d}</td>
    <td style={{...tdS(),textAlign:"left",padding:"4px 8px"}}>
      <div style={{fontWeight:700,color:isLibur?C.gl2:C.wht,fontSize:11}}>{r.dayName}</div>
      <div style={{fontSize:9,color:C.gl2}}>{fDs(r.tgl)}{r.inject?<span style={{color:"#F59E0B",marginLeft:4}}>★</span>:""}{isLibur?<span style={{color:C.gl2,marginLeft:4}}>—</span>:""}</div>
    </td>
    {/* STOK AWAL */}
    {uk.map(s=>[
      <td key={"ai"+s} style={tdS(C.blt)}>{r.awalIsi[s]}</td>,
      <td key={"ak"+s} style={tdS(C.gl2)}>{r.awalTK[s]}</td>
    ])}
    {/* TABUNG MASUK */}
    {uk.map(s=>[
      <td key={"mi"+s} style={tdS(r.masukIsi[s]>0?C.glt:C.gl2,r.masukIsi[s]>0)}>{r.masukIsi[s]||"-"}</td>,
      <td key={"mk"+s} style={tdS(r.masukTK[s]>0?"#F59E0B":C.gl2,r.masukTK[s]>0)}>{r.masukTK[s]||"-"}</td>
    ])}
    {/* TABUNG KELUAR */}
    {uk.map(s=>[
      <td key={"ki"+s} style={tdS(r.keluarIsi[s]>0?C.rlt:C.gl2,r.keluarIsi[s]>0)}>{r.keluarIsi[s]||"-"}</td>,
      <td key={"kk"+s} style={tdS(r.keluarTK[s]!==0?"#F59E0B":C.gl2,r.keluarTK[s]!==0)}>{r.keluarTK[s]!==0?Math.abs(r.keluarTK[s]):"-"}</td>
    ])}
    {/* STOK AKHIR */}
    {uk.map(s=>[
      <td key={"eisi"+s} style={tdS(C.glt,true)}>{r.akhirIsi[s]}</td>,
      <td key={"etk"+s} style={tdS(C.gl2)}>{r.akhirTK[s]}</td>,
      <td key={"etitip"+s} style={tdS(C.blt)}>{r.titipSnap[s]}</td>,
      <td key={"etot"+s} style={tdS(C.olt,true)}>{r.total[s]}</td>
    ])}
    </tr>;
  })}
  </tbody>
  </table>
  <div style={{marginTop:6,fontSize:10,color:C.gl2}}>★ = Inject manual stok awal &nbsp;|&nbsp; — = Tidak ada transaksi &nbsp;|&nbsp; TK = Tabung Kosong</div>
  </div>;
}

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
var[viewTB,setViewTB]=useState(null);// modal lihat
var[editTB,setEditTB]=useState(null);// record yang sedang diedit
useEffect(()=>{if(editTB){setTgl(editTB.tanggal);setCashLaci(String(editTB.cashLaci||""));setRekBSI(String(editTB.rekBSI||""));setRekBCA(String(editTB.rekBCA||""));setPemasukanLain(String(editTB.pemasukanLain||""));}else{setTgl(toDay());setCashLaci("");setRekBSI("");setRekBCA("");}},[editTB]);

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
// DO Gantung: hanya DO yang dibuat pada/sebelum tanggal tutup buku & belum diterima
var doGantung=(data.doList||[]).filter(d=>d.status==="gantung"&&(d.tanggal||"")<=tgl);
var nilaiDOGantung=doGantung.reduce((a,d)=>a+Number(d.totalHPP||0),0);
var doSangkut=(data.doList||[]).filter(d=>d.status==="sangkut"&&(d.tanggal||"")<=tgl);
var nilaiDOSangkut=doSangkut.reduce((a,d)=>a+Number(d.totalHPP||0),0);

// Asset Tabung Milik PT
var titipLuarBal={};
(data.titipList||[]).forEach(t=>{if(t.tipe==="titip_luar"||t.tipe==="tarik_luar"){var m=t.tipe==="titip_luar"?1:-1;(t.items||[]).forEach(it=>{titipLuarBal[it.ukuran]=(titipLuarBal[it.ukuran]||0)+m*Number(it.qty||0);});}});

var hargaTbg=data.company?.hargaTbgKosong||{};
var assetTabungMilikPT=SIZES.reduce((a,s)=>{
  var isiS=(data.stock||{})[s]||0;
  var kosS=getKosong(data,s);
  var titipS=getTitipTotal(data.titipList,s);
  var totalS=isiS+kosS+titipS;
  var titipLuar=Math.max(0,titipLuarBal[s]||0);
  var milikPT=Math.max(0,totalS-titipLuar);
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
var tglTBKemarin=prevTB?.tanggal||null;// tanggal TB terakhir yang dipakai
var deltaSelisih=cashFlowOmset-cashFlowKemarin-(labaBersihH+(Number(pemasukanLain)||0));

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
  id:editTB?editTB.id:uid(),
  tanggal:tgl,
  omzet:omzetH,hpp:hppH,labaKotor:marginH,totalOut:totalOutH,
  pemasukanLain:Number(pemasukanLain)||0,
  labaBersih:labaBersihH+(Number(pemasukanLain)||0),
  cashIn:cashInH,tfIn:tfInH,bonIn:bonInH,
  cashLaci:Number(cashLaci)||0,rekBSI:Number(rekBSI)||0,rekBCA:Number(rekBCA)||0,totalPecah,pecah:{...pecah},
  piutangA,nilaiStokA,pinjamanA,
  nilaiDOGantung,nilaiDOSangkut,
  cashFlowOmset,assetTabungMilikPT,assetArmada,assetValue,
  cashFlowKemarin,deltaSelisih,
  lastEdited:editTB?new Date().toISOString():undefined,
  detail:{penjualan:penjualanDetail,pengeluaran:pengeluaranDetail,doMasuk:doDetail,stokSnapshot:stokSnap,katPengeluaran:katPenH,jumlahTransaksi:penjualanDetail.length,jumlahDO:doDetail.length}
};
if(editTB){
  // Replace record lama
  setData(d=>({...d,tutupBuku:(d.tutupBuku||[]).map(x=>x.id===editTB.id?rec:x)}));
  setEditTB(null);
  toast("✓ Tutup buku diperbarui! Data hari berikutnya otomatis sinkron.");
}else{
  // Hapus record lama dengan tanggal yang sama (hanya 1 per tanggal)
  setData(d=>({...d,tutupBuku:[rec,...(d.tutupBuku||[]).filter(x=>x.tanggal!==tgl)]}));
  toast("✓ Tutup buku harian tersimpan!");
}
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
<Card>
<div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
<Inp label="Tanggal" type="date" value={tgl} onChange={setTgl} style={{maxWidth:220,marginBottom:0}}/>
{editTB&&<div style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:8,padding:"8px 14px",fontSize:12,color:"#FCD34D"}}>
✏️ Mode Edit — {fDs(editTB.tanggal)} · <button onClick={()=>setEditTB(null)} style={{background:"none",border:"none",color:"#FCD34D",cursor:"pointer",fontSize:12,textDecoration:"underline"}}>Batalkan</button>
</div>}
</div>
{editTB&&<div style={{marginTop:8,padding:"8px 12px",background:"#7C2D12",borderRadius:6,fontSize:11,color:"#FCA5A5"}}>⚠️ Mengedit tutup buku akan mempengaruhi perhitungan hari berikutnya. Pastikan data sudah benar sebelum simpan.</div>}
</Card>

{/* 1. CASH WAJIB SETOR KASIR */}
{(()=>{
var allPenTgl=(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var _penjTgl=(data.penjualan||[]).filter(p=>p.tanggal===tgl);
var cashPenjTgl=_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0)+_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="split").reduce((a,p)=>a+Number((p.splitDetail||{}).cash||0),0);
var tfPenjTgl=_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0)+_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="split").reduce((a,p)=>a+Number((p.splitDetail||{}).tf||0),0);
var bonBayarCashTgl=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tgl&&(p.metode||"cash").toLowerCase()==="cash");return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
var bonBayarTFTgl=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tgl&&((p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf"));return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
var penCashTgl=allPenTgl.filter(p=>(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var penTFTgl=allPenTgl.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").reduce((a,p)=>a+Number(p.nominal||0),0);
var totalCashMasuk=cashPenjTgl+bonBayarCashTgl;
var wajibSetorKasir=Math.max(0,totalCashMasuk-penCashTgl);
return <Card style={{border:"2px solid "+C.glt}}>
<div style={{fontWeight:700,color:C.glt,marginBottom:10,fontSize:13}}>🏦 Cash Wajib Setor Kasir — {fDs(tgl)}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
<div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Cash Masuk</div>
{[["Cash Penjualan (semua sales)",cashPenjTgl,C.glt],["Bayar BON Cash (semua)",bonBayarCashTgl,C.glt]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",padding:"5px 8px",background:C.bg,borderRadius:5,marginBottom:3,border:"1px solid "+C.bdr}}>
<span style={{fontSize:11,color:C.gl2}}>{x[0]}</span><b style={{color:x[2],fontSize:12}}>{fR(x[1])}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",padding:"6px 8px",background:C.nav,borderRadius:5,marginBottom:8,border:"1px solid "+C.glt}}>
<b style={{fontSize:12,color:C.wht}}>Total Cash Masuk</b><b style={{color:C.glt,fontSize:13}}>{fR(totalCashMasuk)}</b>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Pengeluaran Cash</div>
{allPenTgl.filter(p=>(p.metode||"cash").toLowerCase()==="cash").map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:C.bg,borderRadius:5,marginBottom:2,border:"1px solid "+C.bdr}}>
<span style={{fontSize:10,color:C.gl2}}>{p.kategori}{p.ket?" — "+p.ket:""} ({p.karyawanNama||"—"})</span><b style={{color:C.rlt,fontSize:11}}>{fR(p.nominal)}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",padding:"6px 8px",background:C.nav,borderRadius:5,border:"1px solid "+C.rlt}}>
<b style={{fontSize:12,color:C.wht}}>Total Pengeluaran Cash</b><b style={{color:C.rlt,fontSize:13}}>{fR(penCashTgl)}</b>
</div>
</div>
<div>
<div style={{background:C.nav,borderRadius:8,padding:14,border:"2px solid "+C.glt,marginBottom:8}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Cash Masuk</div>
<div style={{fontSize:15,fontWeight:900,color:C.glt,marginBottom:8}}>{fR(totalCashMasuk)}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Pengeluaran Cash</div>
<div style={{fontSize:15,fontWeight:900,color:C.rlt,marginBottom:8}}>- {fR(penCashTgl)}</div>
<div style={{height:1,background:C.bdr,marginBottom:8}}/>
<div style={{fontSize:11,color:C.glt,fontWeight:700,marginBottom:4}}>WAJIB SETOR KE BANK</div>
<div style={{fontSize:22,fontWeight:900,color:C.wht,marginBottom:10}}>{fR(wajibSetorKasir)}</div>
<div style={{height:1,background:C.bdr,marginBottom:10}}/>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total Tunai Fisik (pecahan)</div>
<div style={{fontSize:15,fontWeight:700,color:totalPecah>=wajibSetorKasir?C.glt:C.olt,marginBottom:6}}>{fR(totalPecah)}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Selisih Cash</div>
<div style={{fontSize:18,fontWeight:900,color:Math.abs(totalPecah-wajibSetorKasir)<1000?C.glt:C.rlt}}>{fR(totalPecah-wajibSetorKasir)}</div>
</div>
<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid "+C.bdr}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:6,fontWeight:700}}>INFO TRANSFER (masuk rekening langsung)</div>
{[["TF Penjualan",tfPenjTgl],["Bayar BON TF",bonBayarTFTgl],["Pengeluaran TF",penTFTgl]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
<span style={{fontSize:10,color:C.gl2}}>{x[0]}</span><span style={{fontSize:11,color:"#9CA3AF",fontWeight:600}}>{fR(x[1])}</span>
</div>)}
</div>
</div>
</div>
</Card>;
})()}

{/* 2. INPUT CASH FISIK (di dalam CASH FLOW) */}
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📥 Input Cash Fisik:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:8,marginBottom:10}}>
<Inp label="Cash di Laci (Rp)" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Saldo Bank BSI (Rp)" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Saldo Bank BCA (Rp)" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>🪙 Pecahan Kas Fisik (opsional):</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{DENOMS.map(d=><div key={d} style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"5px 11px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}><span style={{color:C.wht,fontSize:13,fontWeight:600}}>{fR(d)}</span><input type="number" value={pecah[d]} placeholder="0" onChange={e=>setPecah(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.wht,fontSize:12,outline:"none",width:74}}/><span style={{color:Number(pecah[d]||0)>0?C.olt:C.gl2,fontWeight:700,fontSize:12}}>{Number(pecah[d]||0)>0?fR(Number(pecah[d])*d):"-"}</span></div>)}
<div style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"9px 11px",background:C.nav,borderTop:"2px solid "+C.bdr}}><b style={{color:C.wht}}>Total Tunai</b><span/><b style={{color:C.glt}}>{fR(totalPecah)}</b></div>
</div>
</Card>

{/* 3. P&L HARI INI */}
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

{/* 4. CASH FLOW / OMSET */}
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
</Card>

{/* 5. VERIFIKASI CASH FLOW */}
<Card style={{border:"1px solid "+(Math.abs(deltaSelisih)<1000?C.glt:C.olt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>✅ Verifikasi Cash Flow</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden"}}>
{[[(tglTBKemarin?"Cash Flow "+fDs(tglTBKemarin):"Belum ada TB sebelumnya"),cashFlowKemarin,C.gl2,false],["Laba Hari Ini",labaBersihH+(Number(pemasukanLain)||0),C.glt,false],["Total Cash Flow Hari Ini",cashFlowOmset,C.blt,true],["SELISIH",deltaSelisih,Math.abs(deltaSelisih)<1000?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span><span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
{Math.abs(deltaSelisih)<1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.grn,borderRadius:6,fontSize:12,fontWeight:700,color:"white"}}>✅ Selisih = 0. Cash flow balance!</div>}
{Math.abs(deltaSelisih)>=1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.rdk,borderRadius:6,fontSize:12,color:"white"}}>⚠️ Ada selisih {fR(Math.abs(deltaSelisih))}. Periksa input cash atau ada transaksi yang terlewat.</div>}
</Card>

{/* 6. TOTAL ASSET */}
<Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>🏦 TOTAL ASSET</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[["Total Cash Flow",cashFlowOmset,C.blt,false],["Asset Tabung Milik PT",assetTabungMilikPT,C.gl2,false],["Asset Armada",assetArmada,C.gl2,false],["TOTAL ASSET",assetValue,C.olt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"12px 14px":"8px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?14:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?800:400}}>{x[0]}</span><span style={{fontSize:x[3]?18:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
</Card>

{/* 7. LAPORAN STOK HARIAN */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 Laporan Stok Harian</div>
{(()=>{
var bulanTgl=tgl.slice(0,7);
var rowsTB=buildStokHarian(data,bulanTgl).filter(r=>r.tgl<=tgl).slice(-3);
var uk=["5.5 kg","12 kg","50 kg"];var ukL=["5,5kg","12kg","50kg"];
return <div style={{overflowX:"auto"}}>
<table style={{borderCollapse:"collapse",width:"100%",fontSize:10}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"left",borderBottom:"1px solid "+C.bdr}}>Hari/Tgl</th>
{uk.map((s,i)=><th key={s} colSpan={4} style={{padding:"5px 6px",color:"white",textAlign:"center",borderBottom:"1px solid "+C.bdr,borderLeft:"2px solid "+C.bdr,background:"#1E3A5F",fontSize:9}}>{ukL[i]}: isi | TK | Titip | Total</th>)}
</tr></thead>
<tbody>
{rowsTB.map((r,i)=><tr key={r.tgl} style={{background:r.tgl===tgl?C.nav:C.bg,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px"}}><div style={{fontWeight:700,color:r.tgl===tgl?C.blt:C.gl2,fontSize:11}}>{r.dayName}</div><div style={{fontSize:9,color:C.gl2}}>{fDs(r.tgl)}</div></td>
{uk.map(s=>[<td key={"i"+s} style={{padding:"4px 6px",textAlign:"center",color:C.glt,fontWeight:700,borderLeft:"2px solid "+C.bdr}}>{r.akhirIsi[s]}</td>,<td key={"k"+s} style={{padding:"4px 6px",textAlign:"center",color:C.gl2}}>{r.akhirTK[s]}</td>,<td key={"t"+s} style={{padding:"4px 6px",textAlign:"center",color:C.blt}}>{r.titipSnap[s]}</td>,<td key={"o"+s} style={{padding:"4px 6px",textAlign:"center",color:C.olt,fontWeight:700}}>{r.total[s]}</td>])}
</tr>)}
</tbody>
</table>
</div>;
})()}
</Card>

{/* 8. REKAP TABUNG */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Rekap Tabung</div>
{(()=>{
var rowsTBRekap=buildStokHarian(data,tgl.slice(0,7)).filter(r=>r.tgl<=tgl);
var tbRow=rowsTBRekap.length>0?rowsTBRekap[rowsTBRekap.length-1]:null;
var titipLuarBalTB={};
(data.titipList||[]).filter(t=>t.tipe==="titip_luar"||t.tipe==="tarik_luar").forEach(t=>{var m=t.tipe==="titip_luar"?1:-1;(t.items||[]).forEach(it=>{titipLuarBalTB[it.ukuran]=(titipLuarBalTB[it.ukuran]||0)+m*Number(it.qty||0);});});
var rows=[
["Di Gudang (Isi)",SIZES.map(s=>tbRow?tbRow.akhirIsi[s]:((data.stock||{})[s]||0)),C.glt,false],
["Kosong di Gudang",SIZES.map(s=>tbRow?tbRow.akhirTK[s]:getKosong(data,s)),C.gl2,false],
["Titip ke Konsumen",SIZES.map(s=>tbRow?tbRow.titipSnap[s]:getTitipTotal(data.titipList,s)),C.blt,false],
["Total Keseluruhan",SIZES.map(s=>tbRow?tbRow.total[s]:0),C.olt,true],
["Titipan Pihak Lain di PT",SIZES.map(s=>Math.max(0,titipLuarBalTB[s]||0)),"#6B7280",false],
["MILIK PT HOE TRANGSA",SIZES.map(s=>Math.max(0,(tbRow?tbRow.total[s]:0)-(titipLuarBalTB[s]||0))),C.wht,true],
];
return <div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.nav}}>
{["Keterangan","5,5 kg","12 kg","50 kg"].map(h=><th key={h} style={{padding:"8px 10px",color:C.gl2,fontWeight:700,textAlign:h==="Keterangan"?"left":"center",borderBottom:"2px solid "+C.bdr}}>{h}</th>)}
</tr></thead>
<tbody>
{rows.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:r[3]?C.nav:"transparent"}}>
<td style={{padding:"7px 10px",color:r[3]?C.wht:C.gl2,fontWeight:r[3]?700:400}}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{padding:"7px 10px",textAlign:"center",color:r[2],fontWeight:r[3]?800:600,fontSize:r[3]?14:12}}>{v}</td>)}
</tr>)}
</tbody>
</table>
</div>;
})()}
</Card>

<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
<Btn color={editTB?"orange":"green"} onClick={saveHarian}>{editTB?"💾 Simpan Perubahan":"💾 Simpan Tutup Buku"}</Btn>
<button onClick={()=>doPrint("_tb_hari")} style={{background:C.blu,color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_tb_hari",makeFileName("tb",tgl,"","png"))} style={{background:"#1D6A96",color:"#fff",border:"none",padding:"9px 14px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>💾 PNG</button>
<button onClick={()=>doCopyPNG("_tb_hari")} style={{background:"#145A32",color:"#fff",border:"none",padding:"9px 14px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>📋 Copy</button>
<span style={{fontSize:10,color:"#888",fontStyle:"italic",alignSelf:"center"}}>💡 PDF: <b>{makeFileName("tb",tgl,"","pdf")}</b></span>
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
  <div style={{display:"flex",gap:4}}>
<button onClick={()=>setViewTB(r)} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 8px",color:C.blt,cursor:"pointer",fontSize:11,fontWeight:700}}>👁️ Lihat</button>
<button onClick={()=>{setEditTB(r);setTab("harian");window.scrollTo(0,0);}} style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"4px 8px",color:"#FCD34D",cursor:"pointer",fontSize:11,fontWeight:700}}>✏️ Edit</button>
<button onClick={()=>{setViewTB(r);setTimeout(()=>doPrint("_tb_view"),300);}} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 8px",color:C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>🖨️ Cetak</button>
</div>
])}/>
</Card>}

{/* Modal Detail Tutup Buku */}
{viewTB&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:9000,overflowY:"auto",padding:16}}>
<div id="_tb_view" style={{maxWidth:680,margin:"0 auto",background:C.card,borderRadius:12,border:"1px solid "+C.bdr,padding:20}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
<div><div style={{fontSize:16,fontWeight:800,color:C.wht}}>📒 Laporan Tutup Buku</div><div style={{fontSize:12,color:C.gl2}}>{viewTB.bulan?BULAN_ID[Number(viewTB.bulan.split("-")[1])-1]+" "+viewTB.bulan.split("-")[0]:fDs(viewTB.tanggal)}</div></div>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"6px 14px",color:C.wht,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>

{/* Ringkasan P&L */}

{/* SECTION 1: CASH FLOW */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.blt,fontSize:12}}>💰 CASH FLOW / OMSET</div>
{[["Cash di Laci",viewTB.cashLaci,C.wht],["Bank BSI",viewTB.rekBSI,C.gl2],["Bank BCA",viewTB.rekBCA,C.gl2],["Total Cash + Bank",(viewTB.cashLaci||0)+(viewTB.rekBSI||0)+(viewTB.rekBCA||0),C.wht,true],["Stok Isi Gudang (HPP)",viewTB.nilaiStokA,C.gl2],["DO Gantung",viewTB.nilaiDOGantung||0,"#F59E0B"],["BON Konsumen",viewTB.piutangA,C.gl2],["Pinjaman Karyawan",viewTB.pinjamanA,C.gl2],["PIUTANG & MODAL BERJALAN",(viewTB.nilaiStokA||0)+(viewTB.nilaiDOGantung||0)+(viewTB.piutangA||0)+(viewTB.pinjamanA||0),C.wht,true],["TOTAL CASH FLOW",viewTB.cashFlowOmset,C.glt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 2: P&L */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.glt,fontSize:12}}>📊 P&L — {fDs(viewTB.tanggal)}</div>
{[["Omzet",viewTB.omzet,C.wht],["HPP / Modal",viewTB.hpp,C.gl2],["Laba Kotor",viewTB.labaKotor,C.blt,true],["Pengeluaran Ops",viewTB.totalOut,C.rlt],["Pemasukan Lainnya",viewTB.pemasukanLain||0,"#F59E0B"],["LABA BERSIH",viewTB.labaBersih,viewTB.labaBersih>=0?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 3: VERIFIKASI */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.blt,fontSize:12}}>✅ Verifikasi Cash Flow</div>
{[["Cash Flow Kemarin",viewTB.cashFlowKemarin||0,C.gl2],["Laba Hari Ini",viewTB.labaBersih,C.glt],["Cash Flow Hari Ini",viewTB.cashFlowOmset,C.blt,true],["Selisih",viewTB.deltaSelisih||0,Math.abs(viewTB.deltaSelisih||0)<1000?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 4: TOTAL ASSET */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.olt,fontSize:12}}>🏦 TOTAL ASSET</div>
{[["Total Cash Flow",viewTB.cashFlowOmset,C.blt],["Asset Tabung Milik PT",viewTB.assetTabungMilikPT||0,C.gl2],["Asset Armada",viewTB.assetArmada||0,C.gl2],["TOTAL ASSET",viewTB.assetValue,C.olt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 5: REKAP TABUNG */}
{viewTB.detail?.stokSnapshot&&<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.gl2,fontSize:12}}>📦 Rekap Tabung</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.card}}>{["Keterangan","5,5 kg","12 kg","50 kg"].map(h=><th key={h} style={{padding:"6px 10px",color:C.gl2,fontWeight:700,textAlign:h==="Keterangan"?"left":"center",borderBottom:"1px solid "+C.bdr}}>{h}</th>)}</tr></thead>
<tbody>
{[["(Tbg+Isi) Di Gudang","isi",C.glt],["Tbg Kosong Di Gudang","kosong",C.gl2],["Titip di Konsumen","titip",C.blt]].map((row,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"6px 10px",color:C.gl2}}>{row[0]}</td>
{SIZES.map((s,j)=><td key={j} style={{padding:"6px 10px",textAlign:"center",color:row[2],fontWeight:600}}>{(viewTB.detail.stokSnapshot[s]||{})[row[1]]||0}</td>)}
</tr>)}
<tr style={{background:C.card,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"7px 10px",color:C.wht,fontWeight:700}}>Total Keseluruhan</td>
{SIZES.map((s,j)=>{var sn=viewTB.detail.stokSnapshot[s]||{};var tot=(sn.isi||0)+(sn.kosong||0)+(sn.titip||0);return <td key={j} style={{padding:"7px 10px",textAlign:"center",color:C.olt,fontWeight:800,fontSize:14}}>{tot}</td>;})}
</tr>
</tbody>
</table>
</div>}

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

<div style={{marginTop:14,display:"flex",gap:8,flexWrap:"wrap"}}>
{(()=>{var fn=makeFileName("tb",viewTB.tanggal,"","");return <>
<button onClick={()=>doPrint("_tb_view",fn+".pdf")} style={{background:C.blu,border:"none",borderRadius:8,padding:"8px 16px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_tb_view",fn+".png")} style={{background:"#1D6A96",border:"none",borderRadius:8,padding:"8px 14px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_tb_view")} style={{background:"#145A32",border:"none",borderRadius:8,padding:"8px 14px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>📋 Copy PNG</button>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",color:C.wht,cursor:"pointer",fontWeight:700,fontSize:12}}>✕ Tutup</button>
</>;})()}
</div>
</div>
</div>}
</div>;
}

// ─── LAPORAN v4 (FilterTbl di semua tab, 2 tab baru) ──────────────────────────
function LaporanMod({data,toast}){
var C=useTheme();
var[mode,setMode]=useState("bulanan");var[bln,setBln]=useState(toMonth());var[tgl,setTgl]=useState(toDay());var[tab,setTab]=useState("harian");
var[tglHarian,setTglHarian]=useState(toDay());
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
// Rekap pengeluaran per kategori bulan ini
var penBlnAll=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var katMap={};
penBlnAll.forEach(p=>{var k=p.kategori||"Lainnya";katMap[k]=(katMap[k]||0)+Number(p.nominal||0);});
var katArr=Object.entries(katMap).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({kategori:k,total:v}));
var totalPenBln=penBlnAll.reduce((a,p)=>a+Number(p.nominal||0),0);

// Tren pengeluaran per bulan (6 bulan terakhir)
var bulanList=[];
var nowM=new Date();for(var mi=5;mi>=0;mi--){var d2=new Date(nowM.getFullYear(),nowM.getMonth()-mi,1);var ym=d2.getFullYear()+"-"+String(d2.getMonth()+1).padStart(2,"0");bulanList.push(ym);}
var trenPen=bulanList.map(ym=>{var total=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(ym)).reduce((a,p)=>a+Number(p.nominal||0),0);return{bulan:ym.slice(5)+"/"+ym.slice(2,4),total};});

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
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>{[["harian","🗓️ Lap. Harian"],["ringkasan","📊 Ringkasan"],["grafik","📈 Grafik"],["stok","📋 Stok Harian"],["pengeluaran","💸 Pengeluaran"],["sales","👤 Per Sales"],["kategori","🏷️ Per Kategori"],["produk","📦 Per Produk"],["pelanggan","👥 Per Pelanggan"],["matrix","📋 Sales×Kategori"],["detail","🔍 Detail"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"6px 11px",fontWeight:700,fontSize:11,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="harian"&&(()=>{
// ── Data hari ini ──
var HARI_ID=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
var dtH=new Date(tglHarian+"T00:00:00");
var hariLabel=HARI_ID[dtH.getDay()]+", "+dtH.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
var penjH=(data.penjualan||[]).filter(p=>p.tanggal===tglHarian);
var doH=(data.doList||[]).filter(d=>d.tanggal===tglHarian&&(d.status||"diterima")==="diterima");
var penH=(data.pengeluaran||[]).filter(p=>p.tanggal===tglHarian);
var bonBayarH=(data.bon||[]).flatMap(b=>(b.pembayaran||[]).filter(px=>px.tanggal===tglHarian).map(px=>({...px,konsumen:b.konsumen,salesId:b.salesId,salesNama:b.salesNama||""})));
// Stok dari buildStokHarian
var bulanH=tglHarian.slice(0,7);
var stokRows=buildStokHarian(data,bulanH).filter(r=>r.tgl<=tglHarian);
var stokRow=stokRows.length>0?stokRows[stokRows.length-1]:null;
// HPP per ukuran (dari modalHistory)
var getHpp=function(uk){var mh=(data.modalHistory||[]).filter(m=>m.ukuran===uk&&m.jenis==="Isi"&&m.tanggal<=tglHarian);return mh.length>0?mh[0].harga:0;};
// Kelompok per sales
var salesGroups={};
penjH.forEach(p=>{var empP=(data.employees||[]).find(e=>e.id===p.salesId);var sNama=empP?.nama||p.salesNama||p.sales||"";var key=sNama||("(Tanpa Sales)");if(!salesGroups[key])salesGroups[key]={nama:key,items:[],omzet:0,margin:0};salesGroups[key].items.push(p);salesGroups[key].omzet+=(p.total||0);salesGroups[key].margin+=(p.margin||0);});
// Pengeluaran per sales
var penPerSales={};
penH.forEach(p=>{var key=p.karyawanNama||"(Umum)";if(!penPerSales[key])penPerSales[key]=0;penPerSales[key]+=Number(p.nominal||0);});
var totalPenH=penH.reduce((a,p)=>a+Number(p.nominal||0),0);
var totalOmzetH=penjH.reduce((a,p)=>a+(p.total||0),0);
var totalMarginH=penjH.reduce((a,p)=>a+(p.margin||0),0);
var totalBonH=bonBayarH.reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
// Print styles
var PS={page:{background:"white",color:"#111",fontFamily:"'Plus Jakarta Sans',Arial,sans-serif",padding:20,maxWidth:900,margin:"0 auto"},h1:{fontSize:18,fontWeight:800,color:"#0a1f44",marginBottom:2},h2:{fontSize:13,fontWeight:700,color:"#0a1f44",margin:"16px 0 8px",borderBottom:"2px solid #0a1f44",paddingBottom:4},th:{background:"#0a1f44",color:"white",padding:"6px 8px",fontSize:10,fontWeight:700,textAlign:"center",border:"1px solid #ccc"},td:{padding:"5px 8px",fontSize:11,border:"1px solid #ddd",verticalAlign:"top"},tdL:{padding:"5px 8px",fontSize:11,border:"1px solid #ddd",textAlign:"left"},tbl:{width:"100%",borderCollapse:"collapse",marginBottom:10},sub:{fontSize:11,color:"#555",marginBottom:2}};

return <div>
<div style={{display:"flex",gap:10,alignItems:"flex-end",marginBottom:14,flexWrap:"wrap"}}>
<Inp label="Tanggal Laporan" type="date" value={tglHarian} onChange={setTglHarian} style={{maxWidth:200,marginBottom:0}}/>
<button onClick={()=>doPrint("_lap_harian")} style={{background:"#0a1f44",color:"white",border:"none",padding:"9px 18px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<span style={{fontSize:10,color:C.gl2,fontStyle:"italic",alignSelf:"center"}}>💡 PDF: Lap-Harian-{tglHarian}.pdf</span>
</div>

<div id="_lap_harian" style={PS.page}>
{/* HEADER */}
<div style={{textAlign:"center",borderBottom:"3px solid #0a1f44",paddingBottom:12,marginBottom:16}}>
<div style={{fontSize:20,fontWeight:900,color:"#0a1f44"}}>LAPORAN HARIAN OPERASIONAL</div>
<div style={{fontSize:14,fontWeight:700,color:"#0a1f44"}}>{(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()}</div>
<div style={{fontSize:12,color:"#555",marginTop:4}}>{hariLabel}</div>
</div>

{/* 1. MUTASI STOK */}
<div style={PS.h2}>1. MUTASI STOK TABUNG</div>
{stokRow?<table style={PS.tbl}>
<thead><tr>
<th style={{...PS.th,textAlign:"left"}}>Keterangan</th>
{SIZES.map(s=>[<th key={"isi"+s} style={PS.th}>isi {s}</th>,<th key={"tk"+s} style={PS.th}>TK {s}</th>])}
{SIZES.map(s=><th key={"tot"+s} style={PS.th}>Total {s}</th>)}
</tr></thead>
<tbody>
{[
["Stok Awal",SIZES.flatMap(s=>[stokRow.awalIsi[s],stokRow.awalTK[s]])],
["Tabung Masuk (DO+Return)",SIZES.flatMap(s=>[stokRow.masukIsi[s],stokRow.masukTK[s]])],
["Tabung Keluar (Penjualan)",SIZES.flatMap(s=>[stokRow.keluarIsi[s],Math.abs(stokRow.keluarTK[s])])],
["Stok Akhir",SIZES.flatMap(s=>[stokRow.akhirIsi[s],stokRow.akhirTK[s]])],
].map((r,i)=><tr key={i} style={{background:i===3?"#EFF6FF":"white"}}>
<td style={{...PS.tdL,fontWeight:i===3?700:400}}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{...PS.td,textAlign:"center",fontWeight:i===3?700:400}}>{v}</td>)}
{i===3?SIZES.map(s=><td key={s} style={{...PS.td,textAlign:"center",fontWeight:700,color:"#1D4ED8"}}>{stokRow.total[s]}</td>):<>{SIZES.map(s=><td key={s} style={{...PS.td,textAlign:"center",color:"#888"}}>—</td>)}</>}
</tr>)}
</tbody>
</table>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada data stok untuk tanggal ini. Gunakan Inject Stok Awal terlebih dahulu.</div>}

{/* 2. DELIVERY ORDER */}
<div style={PS.h2}>2. DELIVERY ORDER MASUK</div>
{doH.length>0?<><table style={PS.tbl}>
<thead><tr>{["Trip","SPPBE","Ukuran","Qty","HPP/Unit","Total HPP","Driver"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{doH.map((d,i)=><tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{d.trip}</td><td style={PS.td}>{d.sppbe}</td><td style={PS.td}>{d.ukuran}</td>
<td style={{...PS.td,textAlign:"center",fontWeight:700}}>{d.qty}</td>
<td style={{...PS.td,textAlign:"right"}}>{fR(d.hppUnit||0)}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(d.totalHPP||0)}</td>
<td style={PS.td}>{d.dibuatOleh||"-"}</td>
</tr>)}
</tbody>
</table>
<div style={{fontSize:11,fontWeight:700,marginBottom:4}}>
Total DO: {SIZES.map(s=>{var q=doH.filter(d=>d.ukuran===s).reduce((a,d)=>a+Number(d.qty||0),0);return q>0?s+" = "+q+" tab":null;}).filter(Boolean).join(", ")} | Total HPP: {fR(doH.reduce((a,d)=>a+Number(d.totalHPP||0),0))}
</div></>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada DO masuk hari ini.</div>}

{/* 3. PENJUALAN PER SALES */}
<div style={PS.h2}>3. PENJUALAN — {penjH.length} Invoice | Omzet: {fR(totalOmzetH)} | Margin: {fR(totalMarginH)}</div>
{Object.values(salesGroups).length>0?Object.values(salesGroups).map((sg,gi)=>{
var q55sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q12sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q50sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var sgCash=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var sgTF=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var sgBon=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
return <div key={gi} style={{marginBottom:12}}>
<div style={{fontWeight:700,fontSize:12,color:"#0a1f44",background:"#EFF6FF",padding:"5px 8px",borderRadius:4,marginBottom:5}}>── {sg.nama} &nbsp;—&nbsp; Omzet: {fR(sg.omzet)} | Margin: {fR(sg.margin)}</div>
<table style={PS.tbl}>
<thead><tr>{["No. Invoice","Konsumen","5,5kg","12kg","50kg","Total","Bayar"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{sg.items.map((p,i)=>{
var q55p=(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var q12p=(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var q50p=(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var byr=(p.bayar||"").toLowerCase();
return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{p.noInv}</td>
<td style={{...PS.tdL,fontWeight:600}}>{p.konsumen}</td>
<td style={{...PS.td,textAlign:"center"}}>{q55p||"—"}</td>
<td style={{...PS.td,textAlign:"center"}}>{q12p||"—"}</td>
<td style={{...PS.td,textAlign:"center"}}>{q50p||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(p.total)}</td>
<td style={{...PS.td,textAlign:"center",color:byr==="cash"?"#15803D":byr==="bon"?"#DC2626":"#1D4ED8",fontWeight:600}}>{p.bayar}</td>
</tr>;})}
<tr style={{background:"#EFF6FF",fontWeight:700}}>
<td colSpan={2} style={{...PS.tdL,fontWeight:700,color:"#0a1f44"}}>Sub-total {sg.nama}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q55sg||"—"}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q12sg||"—"}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q50sg||"—"}</td>
<td colSpan={2} style={{...PS.td,textAlign:"left",color:"#0a1f44"}}>
Cash: {fR(sgCash)} &nbsp;|&nbsp; TF: {fR(sgTF)} &nbsp;|&nbsp; <span style={{color:"#888"}}>BON: {fR(sgBon)}</span> &nbsp;|&nbsp; Total: <b>{fR(sg.omzet)}</b>
</td>
</tr>
</tbody>
</table>
</div>;}):
<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada penjualan hari ini.</div>}
{penjH.length>0&&(()=>{
var totCashH=penjH.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var totTFH=penjH.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var totBonH2=penjH.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
return <table style={{...PS.tbl,marginTop:6}}>
<thead><tr style={{background:"#0a1f44"}}><th colSpan={4} style={{...PS.th,textAlign:"left",fontSize:11}}>TOTAL PENJUALAN — {hariLabel}</th></tr>
<tr style={{background:"#0a1f44"}}>{["Total Cash","Total Transfer","BON / Piutang Baru","TOTAL OMZET"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody><tr style={{background:"#EFF6FF",fontWeight:700}}>
<td style={{...PS.td,textAlign:"right",color:"#15803D"}}>{fR(totCashH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#1D4ED8"}}>{fR(totTFH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#888"}}>{fR(totBonH2)}</td>
<td style={{...PS.td,textAlign:"right",color:"#0a1f44",fontSize:13}}>{fR(totalOmzetH)}</td>
</tr></tbody>
</table>;})()}

{/* 4. PEMBAYARAN BON */}
<div style={PS.h2}>4. PEMBAYARAN BON / PIUTANG</div>
{bonBayarH.length>0?<><table style={PS.tbl}>
<thead><tr>{["Konsumen","Nominal Bayar","Metode","Sales Penerima"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{bonBayarH.map((b,i)=>{var metBon=(b.metode||"cash").toLowerCase();return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.tdL}>{b.konsumen}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(b.jumlah||b.nominal||0)}</td>
<td style={{...PS.td,textAlign:"center",color:metBon==="cash"?"#15803D":"#1D4ED8",fontWeight:600}}>{b.metode||"Cash"}</td>
<td style={PS.td}>{b.salesNama||"-"}</td>
</tr>;})}
</tbody>
</table>
{(()=>{
var bonCashH=bonBayarH.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFH=bonBayarH.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
return <div style={{display:"flex",gap:12,marginTop:6,padding:"6px 10px",background:"#FEF3C7",borderRadius:4,fontSize:11,fontWeight:700}}>
<span>Cash: <b style={{color:"#15803D"}}>{fR(bonCashH)}</b></span>
<span>Transfer: <b style={{color:"#1D4ED8"}}>{fR(bonTFH)}</b></span>
<span style={{marginLeft:"auto"}}>TOTAL: <b style={{color:"#0a1f44"}}>{fR(totalBonH)}</b></span>
</div>;})()} 
</>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada pembayaran BON hari ini.</div>}

{/* 5. PENGELUARAN */}
<div style={PS.h2}>5. PENGELUARAN OPERASIONAL</div>
{penH.length>0?<>{(()=>{
var penCashH=penH.filter(p=>(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var penTFH=penH.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").reduce((a,p)=>a+Number(p.nominal||0),0);
return <>
<table style={PS.tbl}>
<thead><tr>{["Kategori","Keterangan","Atas Nama","Metode","Nominal"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{penH.map((p,i)=>{var met=(p.metode||"cash").toLowerCase();return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{p.kategori}</td>
<td style={PS.tdL}>{p.ket||"-"}</td>
<td style={PS.td}>{p.karyawanNama||"-"}</td>
<td style={{...PS.td,textAlign:"center",color:met==="cash"?"#15803D":"#1D4ED8",fontWeight:600}}>{p.metode||"Cash"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700,color:"#DC2626"}}>{fR(p.nominal)}</td>
</tr>;})}
</tbody>
</table>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:6}}>
<div style={{background:"#FEF2F2",borderRadius:4,padding:"6px 10px",border:"1px solid #FECACA"}}>
<div style={{fontSize:9,color:"#6B7280"}}>Total Cash ({penH.filter(p=>(p.metode||"cash").toLowerCase()==="cash").length} trx)</div>
<div style={{fontWeight:700,fontSize:12,color:"#DC2626"}}>{fR(penCashH)}</div>
</div>
<div style={{background:"#EFF6FF",borderRadius:4,padding:"6px 10px",border:"1px solid #BFDBFE"}}>
<div style={{fontSize:9,color:"#6B7280"}}>Total Transfer ({penH.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").length} trx)</div>
<div style={{fontWeight:700,fontSize:12,color:"#1D4ED8"}}>{fR(penTFH)}</div>
</div>
<div style={{background:"#FEE2E2",borderRadius:4,padding:"6px 10px",border:"1px solid #FCA5A5"}}>
<div style={{fontSize:9,color:"#6B7280"}}>TOTAL PENGELUARAN ({penH.length} trx)</div>
<div style={{fontWeight:700,fontSize:13,color:"#DC2626"}}>{fR(totalPenH)}</div>
</div>
</div>
</>;})()} 
</>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada pengeluaran hari ini.</div>}

{/* 6. RINGKASAN PER SALES */}
<div style={PS.h2}>6. RINGKASAN PER SALES</div>
<div style={{overflowX:"auto"}}>
<table style={{...PS.tbl,minWidth:700}}>
<thead>
<tr style={{background:"#0a1f44"}}>
<th rowSpan={2} style={{...PS.th,textAlign:"left",fontSize:8}}>Sales</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>5,5 kg</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>12 kg</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>50 kg</th>
<th colSpan={3} style={{...PS.th,fontSize:8}}>Penjualan</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>Bayar BON</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Pemasukan Kas</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Margin</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Pengeluaran</th>
</tr>
<tr style={{background:"#0a1f44"}}>
{["Qty","Nominal","Qty","Nominal","Qty","Nominal","Cash","TF","BON*","Cash","TF"].map(h=><th key={h} style={{...PS.th,fontSize:7,padding:"3px 4px"}}>{h}</th>)}
</tr>
</thead>
<tbody>
{Object.values(salesGroups).map((sg,i)=>{
var q55s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q12s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q50s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var nom55s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var nom12s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var nom50s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var sgCashS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var sgTFS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var sgBonS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
var bonCashS=bonBayarH.filter(b=>b.salesNama===sg.nama&&(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFS=bonBayarH.filter(b=>b.salesNama===sg.nama&&((b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf")).reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var pemasukanKasS=sgCashS+sgTFS+bonCashS+bonTFS;
var penSales=(penPerSales[sg.nama]||0);
return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={{...PS.tdL,fontWeight:700,fontSize:9,padding:"3px 5px"}}>{sg.nama}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q55s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom55s?fR(nom55s):"—"}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q12s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom12s?fR(nom12s):"—"}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q50s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom50s?fR(nom50s):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{fR(sgCashS)}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{fR(sgTFS)}</td>
<td style={{...PS.td,textAlign:"right",color:"#888",fontSize:9,padding:"3px 4px"}}>{sgBonS?fR(sgBonS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{bonCashS?fR(bonCashS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{bonTFS?fR(bonTFS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700,color:"#0a1f44",fontSize:9,padding:"3px 4px"}}>{fR(pemasukanKasS)}</td>
<td style={{...PS.td,textAlign:"right",color:"#15803D",fontWeight:700,fontSize:9,padding:"3px 4px"}}>{fR(sg.margin)}</td>
<td style={{...PS.td,textAlign:"right",color:"#DC2626",fontSize:9,padding:"3px 4px"}}>{penSales?fR(penSales):"—"}</td>
</tr>;})}
{(()=>{
var tq55=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tq12=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tq50=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tCash=penjH.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var tTF=penjH.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var tBonPenj=penjH.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
var bonCashTotal=bonBayarH.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFTotal=bonBayarH.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var totalPemasukan=tCash+tTF+bonCashTotal+bonTFTotal;
return <tr style={{background:"#0a1f44",color:"white",fontWeight:700}}>
<td style={{...PS.tdL,color:"white",fontSize:9,padding:"4px 5px"}}>TOTAL</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq55||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq12||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq50||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(tCash)}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(tTF)}</td>
<td style={{...PS.td,textAlign:"right",color:"#aaa",fontSize:9,padding:"4px"}}>{tBonPenj?fR(tBonPenj):"—"}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(bonCashTotal)}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(bonTFTotal)}</td>
<td style={{...PS.td,textAlign:"right",color:"#86EFAC",fontSize:10,padding:"4px",fontWeight:700}}>{fR(totalPemasukan)}</td>
<td style={{...PS.td,textAlign:"right",color:"#86EFAC",fontSize:9,padding:"4px"}}>{fR(totalMarginH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#FCA5A5",fontSize:9,padding:"4px"}}>{fR(totalPenH)}</td>
</tr>;})()} 
</tbody>
</table>
</div>
<div style={{fontSize:9,color:"#888",marginTop:4}}>* BON = piutang baru, belum masuk kas &nbsp;|&nbsp; Pemasukan Kas = Cash Penj + TF Penj + Bayar BON Cash + Bayar BON TF</div>

{/* 7. REKAP TABUNG AKHIR */}
<div style={PS.h2}>7. REKAP TABUNG AKHIR HARI</div>
{stokRow?<table style={PS.tbl}>
<thead><tr>{["Keterangan","5,5 kg - Isi","5,5 kg - TK","12 kg - Isi","12 kg - TK","50 kg - Isi","50 kg - TK"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{[["Stok Akhir Isi & TK",SIZES.flatMap(s=>[stokRow.akhirIsi[s],stokRow.akhirTK[s]])],
["Titip di Konsumen",SIZES.flatMap(s=>[stokRow.titipSnap[s],"—"])],
["TOTAL KESELURUHAN",SIZES.flatMap(s=>[stokRow.total[s],"="+stokRow.total[s]])]].map((r,i)=><tr key={i} style={{background:i===2?"#DBEAFE":"white",fontWeight:i===2?700:400}}>
<td style={PS.tdL}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{...PS.td,textAlign:"center"}}>{v}</td>)}
</tr>)}
</tbody>
</table>:<div style={{color:"#888",fontSize:11}}>Data stok tidak tersedia.</div>}

<div style={{marginTop:20,borderTop:"2px solid #0a1f44",paddingTop:10,display:"flex",justifyContent:"space-between",fontSize:10,color:"#555"}}>
<span>Dicetak: {new Date().toLocaleString("id-ID")}</span>
<span>{data.company?.nama||"PT. HOE TRANG SA"}</span>
</div>
</div>
</div>;
})()}

{tab==="ringkasan"&&<>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💰 Komposisi Pembayaran</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>{[["Cash",cash,C.glt],["Transfer",tf,C.blt],["BON",bon,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr,textAlign:"center"}}><div style={{fontSize:11,color:C.gl2}}>{x[0]}</div><div style={{fontSize:14,fontWeight:900,color:x[2]}}>{fR(x[1])}</div><div style={{fontSize:10,color:C.gl2,marginTop:2}}>{omzet>0?(x[1]/omzet*100).toFixed(1):0}%</div></div>)}</div></Card>

</>}
{tab==="sales"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👤 Per Sales</div><FilterTbl columns={salesCols} data={salesArr} empty="Tidak ada data"/></Card>}
{tab==="kategori"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🏷️ Per Kategori Pelanggan</div><FilterTbl columns={katCols} data={katArr} empty="Tidak ada data"/>{katArr.length>0&&<div style={{marginTop:14}}><ResponsiveContainer width="100%" height={220}><PieChart><Pie data={katArr.slice(0,8)} dataKey="omzet" nameKey="kategori" cx="50%" cy="50%" outerRadius={80} label={x=>x.kategori}>{katArr.slice(0,8).map((e,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}</Pie><Tooltip formatter={v=>fR(v)} contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht}}/></PieChart></ResponsiveContainer></div>}</Card>}
{tab==="produk"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Per Produk</div><RTbl headers={["Ukuran","Qty","Omzet","% Omzet"]} rows={prodArr.map(p=>[<b style={{color:C.wht}}>{p.ukuran}</b>,<b style={{color:C.glt}}>{p.qty} tab</b>,<b style={{color:C.blt}}>{fR(p.omzet)}</b>,(omzet>0?(p.omzet/omzet*100).toFixed(1):0)+"%"])}/></Card>}
{tab==="pelanggan"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👥 Ranking Pelanggan per Omzet</div><FilterTbl columns={plgCols} data={plgArr} empty="Tidak ada data"/></Card>}
{tab==="matrix"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Sales × Kategori</div><FilterTbl columns={skCols} data={skArr} empty="Tidak ada data"/></Card>}
{tab==="detail"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔍 Detail Penjualan ({penjFilt.length})</div><FilterTbl columns={detCols} data={detailRows} empty="Tidak ada data" maxRows={300}/></Card>}

{tab==="stok"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📋 Laporan Stok Harian — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
<TabelStokBulanan data={data} bulan={bln}/>
</Card>
</div>}
{tab==="pengeluaran"&&<div>
{/* Summary */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:8,marginBottom:12}}>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Total Pengeluaran</div><div style={{fontSize:15,fontWeight:900,color:C.rlt}}>{fR(totalPenBln)}</div></div>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Jumlah Kategori</div><div style={{fontSize:15,fontWeight:900,color:C.wht}}>{katArr.length}</div></div>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Jumlah Transaksi</div><div style={{fontSize:15,fontWeight:900,color:C.wht}}>{penBlnAll.length}</div></div>
</div>

{/* Tabel per kategori */}
{katArr.length>0?<Card style={{marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💸 Total per Kategori — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
<RTbl headers={["Kategori","Total","% dari Total","Transaksi"]} rows={katArr.map(k=>{
var pct=totalPenBln>0?Math.round(k.total/totalPenBln*100):0;
var cnt=penBlnAll.filter(p=>(p.kategori||"Lainnya")===k.kategori).length;
return[
<b style={{color:C.wht}}>{k.kategori}</b>,
<b style={{color:C.rlt}}>{fR(k.total)}</b>,
<div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:Math.max(4,pct)+"px",maxWidth:80,height:6,background:C.rlt,borderRadius:3}}/><span style={{color:C.gl2,fontSize:11}}>{pct}%</span></div>,
<span style={{color:C.gl2}}>{cnt}x</span>
];})}/>
</Card>:<Card><div style={{color:C.gl2,textAlign:"center",padding:20}}>Tidak ada pengeluaran bulan ini</div></Card>}

{/* Grafik bar per kategori */}
{katArr.length>0&&<Card style={{marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📊 Grafik per Kategori</div>
<ResponsiveContainer width="100%" height={Math.max(200,katArr.length*35)}><BarChart layout="vertical" data={katArr} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="kategori" stroke={C.gl2} fontSize={9} width={110}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="total" fill="#EF4444" name="Total" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer>
</Card>}

{/* Tren 6 bulan */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📈 Tren Pengeluaran 6 Bulan Terakhir</div>
<ResponsiveContainer width="100%" height={180}><AreaChart data={trenPen} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="bulan" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="total" stroke="#EF4444" fill="#EF4444" fillOpacity={0.25} name="Total Pengeluaran"/>
</AreaChart></ResponsiveContainer>
</Card>

{/* Detail transaksi */}
{penBlnAll.length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Detail Pengeluaran</div>
<RTbl headers={["Tgl","Kategori","Keterangan","Karyawan","Nominal"]} rows={penBlnAll.slice().sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")).slice(0,100).map(p=>[
fDs(p.tanggal),
<Bdg color="red">{p.kategori}</Bdg>,
<span style={{color:C.gl2,fontSize:11}}>{p.ket||"-"}</span>,
p.karyawanNama||"-",
<b style={{color:C.rlt}}>{fR(p.nominal)}</b>
])}/>
</Card>}
</div>}

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

var LOGO_HTS="data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA...";

function KasBankMod({data,setData,user,toast}){
var C=useTheme();
var[tabK,setTabK]=useState("saldo");
var[bulanSaldo,setBulanSaldo]=useState(toMonth());
// Setup saldo awal
var[setupBSI,setSetupBSI]=useState({nominal:String((data.saldoAwalBank||{}).BSI?.nominal||""),tanggal:(data.saldoAwalBank||{}).BSI?.tanggal||""});
var[setupBCA,setSetupBCA]=useState({nominal:String((data.saldoAwalBank||{}).BCA?.nominal||""),tanggal:(data.saldoAwalBank||{}).BCA?.tanggal||""});
// Setoran cash states
var[bulanSetor,setBulanSetor]=useState(toMonth());
var[tglPilih,setTglPilih]=useState([]);// array tgl yang dipilih
var[bankSetor,setBankSetor]=useState("BSI");
var[penyetor,setPenyetor]=useState("Muhammad Haekal");
var[pecahInput,setPecahInput]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[showSlip,setShowSlip]=useState(false);
var[editSetor,setEditSetor]=useState(null);
var[editSetorF,setEditSetorF]=useState(null);

// ── Kalkulasi saldo bank ──
function getMutasiBank(bank){
  var list=[];
  // Saldo awal
  var sa=data.saldoAwalBank?.[bank]||{};
  var saldoAwal=Number(sa.nominal||0);
  var tglAwal=sa.tanggal||"";
  // TF Penjualan masuk ke bank ini
  (data.penjualan||[]).filter(p=>(!tglAwal||p.tanggal>=tglAwal)).forEach(p=>{
    var byr=(p.bayar||"").toLowerCase();
    if(byr==="transfer"&&(p.bank||"BSI")===bank){list.push({tanggal:p.tanggal,ket:"TF Penjualan — "+p.konsumen+" ("+p.noInv+")",jenis:"Masuk TF",masuk:p.total||0,keluar:0,bank});}
    else if(byr==="split"&&Number((p.splitDetail||{}).tf||0)>0&&(p.splitBank||"BSI")===bank){list.push({tanggal:p.tanggal,ket:"TF Penjualan (Split) — "+p.konsumen+" ("+p.noInv+")",jenis:"Masuk TF",masuk:Number(p.splitDetail.tf),keluar:0,bank});}
  });
  // Bayar BON TF masuk ke bank ini
  (data.bon||[]).forEach(b=>(b.pembayaran||[]).filter(px=>((px.metode||"").toLowerCase()==="transfer"||(px.metode||"").toLowerCase()==="tf")&&(px.bank||"BSI")===bank&&(!tglAwal||px.tanggal>=tglAwal)).forEach(px=>{
    list.push({tanggal:px.tanggal,ket:"Bayar BON — "+b.konsumen,jenis:"Masuk TF",masuk:Number(px.jumlah||px.nominal||0),keluar:0,bank});
  }));
  // Pengeluaran TF keluar dari bank ini
  (data.pengeluaran||[]).filter(p=>(p.metode||"").toLowerCase()==="transfer"&&(!tglAwal||p.tanggal>=tglAwal)).forEach(p=>{
    list.push({tanggal:p.tanggal,ket:"Pengeluaran — "+p.kategori+" ("+( p.ket||"")+")",jenis:"Keluar TF",masuk:0,keluar:Number(p.nominal||0),bank});
  });
  // DO selalu potong BSI
  if(bank==="BSI"){
    (data.doList||[]).filter(d=>(d.status||"diterima")==="diterima"&&(!tglAwal||d.tanggal>=tglAwal)).forEach(d=>{
      list.push({tanggal:d.tanggal,ket:"DO — "+d.sppbe+" ("+d.trip+")",jenis:"DO Bayar",masuk:0,keluar:Number(d.totalHPP||0),bank:"BSI"});
    });
  }
  // Setoran cash ke bank ini
  (data.setoranBank||[]).filter(s=>s.bank===bank&&(!tglAwal||s.tanggal>=tglAwal)).forEach(s=>{
    list.push({tanggal:s.tanggal,ket:"Setoran Cash ("+s.tglList?.join(",")+") ",jenis:"Setor Cash",masuk:s.nominal||0,keluar:0,bank});
  });
  // Sort by tanggal
  list.sort((a,b)=>a.tanggal.localeCompare(b.tanggal));
  // Hitung running saldo
  var saldo=saldoAwal;
  list=list.map(x=>{saldo+=x.masuk-x.keluar;return{...x,saldo};});
  return{saldoAwal,list,saldoAkhir:saldo,tglAwal};
}

var mutBSI=getMutasiBank("BSI");
var mutBCA=getMutasiBank("BCA");
var totalBank=mutBSI.saldoAkhir+mutBCA.saldoAkhir;

// ── Kalkulasi per hari untuk kalender setoran ──
var dim=daysInMonth(bulanSetor);
var hariData={};
for(var di=1;di<=dim;di++){
  var tglD=bulanSetor+"-"+String(di).padStart(2,"0");
  // Cek ada tutup buku dengan pecahan
  var tb=(data.tutupBuku||[]).find(t=>t.tanggal===tglD);
  var hasPecah=tb&&(
  (typeof tb.totalPecah==="number"&&tb.totalPecah>0)||
  (typeof tb.totalPecah==="object"&&tb.totalPecah!==null&&Object.values(tb.totalPecah).some(v=>v>0))
);
  var wajibSetor=0;
  if(tb){
    var cashPenj=(data.penjualan||[]).filter(p=>p.tanggal===tglD&&(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
    var bonCashD=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tglD&&(p.metode||"cash").toLowerCase()==="cash");return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
    var penCashD=(data.pengeluaran||[]).filter(p=>p.tanggal===tglD&&(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
    wajibSetor=Math.max(0,cashPenj+bonCashD-penCashD);
  }
  var disetor=(data.kas||{})[tglD]?.disetor;
  // pecah object: dari field pecah (DENOMS qty) atau estimasi dari totalPecah number
var pecahObj={};
if(tb?.pecah&&typeof tb.pecah==="object"){pecahObj=tb.pecah;}
else if(typeof tb?.totalPecah==="number"&&tb.totalPecah>0){
  // estimasi dari total — bagi ke pecahan besar
  var sisa=tb.totalPecah;
  DENOMS.forEach(d=>{var lbr=Math.floor(sisa/d);pecahObj[d]=lbr;sisa-=lbr*d;});
}
hariData[tglD]={hasTB:!!tb,hasPecah,wajibSetor,disetor,pecah:pecahObj,totalPecah:typeof tb?.totalPecah==="number"?tb.totalPecah:0,tglD};
}

// Total dari tgl yang dipilih
var totalPilih=tglPilih.reduce((a,tgl)=>a+(hariData[tgl]?.wajibSetor||0),0);
// Gabungkan pecahan dari tgl yang dipilih
var pecahGabung={};
DENOMS.forEach(d=>{pecahGabung[d]=tglPilih.reduce((a,tgl)=>a+Number((hariData[tgl]?.pecah||{})[d]||0),0);});
// Init pecahInput dari gabungan saat tglPilih berubah
var totalFisik=DENOMS.reduce((a,d)=>a+Number(pecahInput[d]||0)*d,0);
var selisihPecah=totalFisik-totalPilih;

function konfirmasiSetor(){
  if(tglPilih.length===0){toast("Pilih minimal 1 tanggal");return;}
  var newKas={...(data.kas||{})};
  tglPilih.forEach(tgl=>{newKas[tgl]={disetor:true,bank:bankSetor,nominal:hariData[tgl]?.wajibSetor||0,tglSetor:toDay()};});
  var logSetor={id:uid(),tanggal:toDay(),bank:bankSetor,nominal:totalPilih,tglList:tglPilih,penyetor,pecah:{...pecahGabung},pecahReal:{...pecahInput}};
  setData(d=>({...d,kas:newKas,setoranBank:[logSetor,...(d.setoranBank||[])]}));
  setShowSlip(true);
  setTglPilih([]);
  setPecahInput(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
  toast("✓ Setoran Rp "+fR(totalPilih)+" ke "+bankSetor+" berhasil dikonfirmasi!");
}

// Hari ini
var HARI_ID=["Min","Sen","Sel","Rab","Kam","Jum","Sab"];
var firstDay=new Date(bulanSetor+"-01T00:00:00").getDay();
var BULAN_LABEL=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
var blnIdx=Number(bulanSetor.split("-")[1])-1;
var thnIdx=bulanSetor.split("-")[0];

return <div>
<STitle icon="🏦" children="Kas & Bank"/>
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>
{[["saldo","📊 Saldo & Mutasi"],["setor","💵 Setoran Cash"],["setup","⚙️ Setup Saldo Awal"]].map(x=><button key={x[0]} onClick={()=>setTabK(x[0])} style={{background:tabK===x[0]?C.blu:C.nav,color:tabK===x[0]?"white":C.wht,border:"1px solid "+(tabK===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>

{/* ── TAB SALDO & MUTASI ── */}
{tabK==="saldo"&&<div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:14}}>
{[["🏦 Saldo BSI",mutBSI.saldoAkhir,C.blt],["🏦 Saldo BCA",mutBCA.saldoAkhir,"#6B7280"],["💰 Total Saldo Bank",totalBank,C.glt]].map(x=><Card key={x[0]} style={{marginBottom:0,textAlign:"center"}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>{x[0]}</div>
<div style={{fontSize:18,fontWeight:900,color:x[2]}}>{fR(x[1])}</div>
</Card>)}
</div>
{["BSI","BCA"].map(bank=>{
var mut=bank==="BSI"?mutBSI:mutBCA;
return <Card key={bank} style={{marginBottom:10}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:8,fontSize:13}}>🏦 Mutasi Bank {bank} — Saldo Awal ({mut.tglAwal?fDs(mut.tglAwal):"belum diset"}): {fR(mut.saldoAwal)}</div>
{mut.list.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada mutasi</div>:
<div style={{overflowX:"auto"}}>
<RTbl headers={["Tgl","Keterangan","Jenis","Masuk","Keluar","Saldo"]} rows={mut.list.slice().reverse().slice(0,50).map(r=>[
fDs(r.tanggal),
<span style={{fontSize:11}}>{r.ket}</span>,
<Bdg color={r.jenis.includes("Masuk")||r.jenis.includes("Setor")?"green":"red"}>{r.jenis}</Bdg>,
r.masuk>0?<b style={{color:C.glt}}>{fR(r.masuk)}</b>:"—",
r.keluar>0?<b style={{color:C.rlt}}>{fR(r.keluar)}</b>:"—",
<b style={{color:r.saldo>=0?C.wht:C.rlt}}>{fR(r.saldo)}</b>
])}/>
</div>}
</Card>;})}
</div>}

{/* ── TAB SETORAN CASH ── */}
{tabK==="setor"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💵 Pilih Tanggal Setoran</div>
<div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10}}>
<Inp label="Bulan" type="month" value={bulanSetor} onChange={v=>{setBulanSetor(v);setTglPilih([]);}} style={{maxWidth:160,marginBottom:0}}/>
<div style={{fontSize:11,color:C.gl2}}>
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:"#F59E0B",marginRight:4}}/>Tersedia &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.blt,marginRight:4}}/>Dipilih &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.glt,marginRight:4}}/>Sudah Setor &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.bdr,marginRight:4}}/>Tidak Ada
</div>
</div>
{/* Kalender */}
<div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4,marginBottom:12}}>
{HARI_ID.map(h=><div key={h} style={{textAlign:"center",fontSize:10,color:C.gl2,fontWeight:700,padding:"4px 0"}}>{h}</div>)}
{Array.from({length:firstDay}).map((_,i)=><div key={"e"+i}/>)}
{Array.from({length:dim}).map((_,i)=>{
var d2=i+1;
var tglD2=bulanSetor+"-"+String(d2).padStart(2,"0");
var hd=hariData[tglD2]||{};
var isPilih=tglPilih.includes(tglD2);
var bg=hd.disetor?C.grn:isPilih?C.blt:hd.hasPecah?"#F59E0B":C.bdr;
var color=hd.disetor||isPilih||hd.hasPecah?"white":C.gl2;
var canClick=hd.hasPecah&&!hd.disetor;
return <div key={d2} onClick={()=>{if(!canClick)return;setTglPilih(prev=>prev.includes(tglD2)?prev.filter(x=>x!==tglD2):[...prev,tglD2]);setPecahInput(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});}} style={{textAlign:"center",padding:"8px 4px",borderRadius:6,background:bg,color,fontSize:12,fontWeight:isPilih||hd.disetor?700:400,cursor:canClick?"pointer":"default",opacity:hd.hasTB&&!hd.hasPecah?.6:1,border:isPilih?"2px solid white":"2px solid transparent"}}>
<div>{d2}</div>
{hd.wajibSetor>0&&<div style={{fontSize:8,opacity:.85}}>{(hd.wajibSetor/1000).toFixed(0)}k</div>}
</div>;})}
</div>
</Card>

{/* Tabel 1: Rincian per tgl - selalu tampil Opsi A */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📋 Tabel 1 — Rincian Pecahan per Tanggal (dari Tutup Buku)</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"6px 10px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:10}}>Pecahan</th>
{tglPilih.length>0?tglPilih.sort().map(tgl=><th key={tgl} style={{padding:"6px 8px",color:C.blt,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}}>{fDs(tgl)}</th>):<th style={{padding:"6px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10,fontStyle:"italic"}}>(pilih tanggal)</th>}
<th style={{padding:"6px 10px",color:C.olt,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10,fontWeight:700}}>TOTAL</th>
</tr></thead>
<tbody>
{DENOMS.map(d=>{
var totD=tglPilih.reduce((a,tgl)=>a+Number((hariData[tgl]?.pecah||{})[d]||0),0);
return <tr key={d} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"5px 10px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{fR(d)}</td>
{tglPilih.length>0?tglPilih.sort().map(tgl=>{var v=Number((hariData[tgl]?.pecah||{})[d]||0);return <td key={tgl} style={{padding:"5px 8px",textAlign:"center",color:v>0?C.wht:C.gl2,border:"1px solid "+C.bdr}}>{v||"—"}</td>;}):
<td style={{padding:"5px 8px",textAlign:"center",color:C.gl2,border:"1px solid "+C.bdr}}>—</td>}
<td style={{padding:"5px 10px",textAlign:"center",fontWeight:700,color:totD>0?C.olt:C.gl2,border:"1px solid "+C.bdr}}>{totD||"—"}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td style={{padding:"6px 10px",color:C.wht,border:"1px solid "+C.bdr}}>TOTAL</td>
{tglPilih.length>0?tglPilih.sort().map(tgl=><td key={tgl} style={{padding:"6px 8px",textAlign:"center",color:C.glt,border:"1px solid "+C.bdr}}>{fR(hariData[tgl]?.wajibSetor||0)}</td>):
<td style={{padding:"6px 8px",textAlign:"center",color:C.gl2,border:"1px solid "+C.bdr}}>—</td>}
<td style={{padding:"6px 10px",textAlign:"center",fontWeight:900,color:tglPilih.length>0?C.glt:C.gl2,fontSize:14,border:"1px solid "+C.bdr}}>{tglPilih.length>0?fR(totalPilih):"—"}</td>
</tr>
</tbody>
</table>
</div>
</Card>

{/* Tabel 2: selalu tampil */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>✏️ Tabel 2 — Input Real Pecahan (bisa diedit)</div>
{tglPilih.length===0&&<div style={{fontSize:11,color:C.gl2,marginBottom:8,fontStyle:"italic"}}>Pilih tanggal di kalender → Tabel 1 akan muncul dan Tabel 2 terisi otomatis.</div>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
<div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>{tglPilih.length>0?"Terisi dari Tabel 1 · edit sesuai fisik yang ada:":"Input manual pecahan:"}</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
{["Pecahan","T1 (sistem)","T2 (real)","Selisih"].map(h=><th key={h} style={{padding:"5px 8px",color:C.gl2,textAlign:h==="Pecahan"?"left":"center",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{DENOMS.map(d=>{
var sys=pecahGabung[d]||0;
var real=Number(pecahInput[d]||0);
var sel=real-sys;
return <tr key={d} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{fR(d)}</td>
<td style={{padding:"4px 8px",textAlign:"center",color:sys>0?C.blt:C.gl2,border:"1px solid "+C.bdr}}>{sys||"—"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr}}>
<input type="number" value={pecahInput[d]} placeholder={String(sys||0)} onChange={e=>setPecahInput(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"none",borderRadius:4,padding:"3px 6px",color:C.wht,fontSize:11,outline:"none",width:"100%",textAlign:"center"}}/>
</td>
<td style={{padding:"4px 8px",textAlign:"center",color:sel>0?C.glt:sel<0?C.rlt:C.gl2,fontWeight:sel!==0?700:400,border:"1px solid "+C.bdr}}>{tglPilih.length>0?(sel>0?"+"+sel:sel||"—"):"—"}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td colSpan={2} style={{padding:"6px 8px",color:C.wht,border:"1px solid "+C.bdr}}>Total T2 (real)</td>
<td style={{padding:"6px 8px",textAlign:"center",color:C.glt,fontWeight:900,fontSize:13,border:"1px solid "+C.bdr}}>{totalFisik>0?fR(totalFisik):"—"}</td>
<td style={{padding:"6px 8px",textAlign:"center",color:Math.abs(selisihPecah)<1000&&totalFisik>0?C.glt:C.rlt,fontWeight:700,border:"1px solid "+C.bdr}}>{tglPilih.length>0&&totalFisik>0?(selisihPecah>=0?"+":"")+fR(selisihPecah):"—"}</td>
</tr>
</tbody>
</table>
</div>
<div>
<div style={{background:C.nav,borderRadius:8,padding:14,border:"1px solid "+C.bdr,marginBottom:8}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total dari Tabel 1 (sistem)</div>
<div style={{fontSize:16,fontWeight:700,color:C.blt,marginBottom:10}}>{tglPilih.length>0?fR(totalPilih):"—"}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total dari Tabel 2 (real)</div>
<div style={{fontSize:16,fontWeight:700,color:C.glt,marginBottom:10}}>{totalFisik>0?fR(totalFisik):"—"}</div>
<div style={{height:1,background:C.bdr,marginBottom:10}}/>
<div style={{fontSize:11,color:Math.abs(selisihPecah)<1000?C.glt:C.olt,fontWeight:700,marginBottom:4}}>Selisih T1 vs T2</div>
<div style={{fontSize:18,fontWeight:900,color:Math.abs(selisihPecah)<1000&&tglPilih.length>0?C.glt:C.rlt}}>{tglPilih.length>0&&totalFisik>0?(selisihPecah>=0?"+":"")+fR(selisihPecah):"—"}</div>
</div>
<div style={{display:"flex",flexDirection:"column",gap:8}}>
<Sel label="Setor ke Bank" value={bankSetor} onChange={setBankSetor} opts={[{v:"BSI",l:"BSI — Bank Syariah Indonesia"},{v:"BCA",l:"BCA"}]}/>
<Inp label="Penyetor" value={penyetor} onChange={setPenyetor}/>
<Btn color="green" onClick={konfirmasiSetor}>✓ Konfirmasi Setoran {tglPilih.length>0?fR(totalPilih):"—"} → {bankSetor}</Btn>
</div>
</div>
</div>
</Card>

{/* Riwayat setoran bank */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📋 Riwayat Setoran Bank</div>
{(data.setoranBank||[]).length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada setoran</div>:
<div>
{(data.setoranBank||[]).slice(0,20).map((r,ri)=><div key={r.id||ri} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:ri%2===0?C.nav:C.bg,borderRadius:6,marginBottom:4,border:"1px solid "+C.bdr}}>
<div>
<div style={{fontSize:12,fontWeight:700,color:C.wht}}>{fDs(r.tanggal)} — <Bdg color="blue">{r.bank}</Bdg> — <b style={{color:C.glt}}>{fR(r.nominal)}</b></div>
<div style={{fontSize:10,color:C.gl2,marginTop:2}}>Penyetor: {r.penyetor} · Tgl: {r.tglList?.map(t=>fDs(t)).join(", ")}</div>
</div>
<div style={{display:"flex",gap:5}}>
<button onClick={()=>{setEditSetor(r);setShowSlip(true);}} style={{background:C.nav,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 8px",color:C.blt,cursor:"pointer",fontSize:11,fontWeight:700}}>👁️ View</button>
<button onClick={()=>{setEditSetor(r);setEditSetorF({...r});}} style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"4px 8px",color:"#FCD34D",cursor:"pointer",fontSize:11,fontWeight:700}}>✏️ Edit</button>
<button onClick={()=>{setEditSetor(r);setShowSlip(true);setTimeout(()=>doPrint("_slip_setor"),400);}} style={{background:C.nav,border:"1px solid "+C.gl2,borderRadius:6,padding:"4px 8px",color:C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>🖨️ Cetak</button>
</div>
</div>)}
</div>}
</Card>
</div>}

{/* ── TAB SETUP SALDO AWAL ── */}
{tabK==="setup"&&<div>
{["BSI","BCA"].map(bank=>{
var st=bank==="BSI"?setupBSI:setupBCA;
var setSt=bank==="BSI"?setSetupBSI:setSetupBCA;
return <Card key={bank}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>🏦 Bank {bank}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
<Inp label="Tanggal Efektif" type="date" value={st.tanggal} onChange={v=>setSt(p=>({...p,tanggal:v}))}/>
<Inp label="Saldo Awal (Rp)" type="number" value={st.nominal} onChange={v=>setSt(p=>({...p,nominal:v}))} placeholder="0"/>
</div>
<Btn color="blue" onClick={()=>{
var newSaldo={...(data.saldoAwalBank||{})};
newSaldo[bank]={nominal:Number(st.nominal||0),tanggal:st.tanggal};
setData(d=>({...d,saldoAwalBank:newSaldo}));
toast("✓ Saldo awal "+bank+" disimpan!");
}}>💾 Simpan Saldo Awal {bank}</Btn>
<div style={{marginTop:8,fontSize:11,color:C.gl2}}>Saldo saat ini: <b style={{color:C.glt}}>{fR(bank==="BSI"?mutBSI.saldoAkhir:mutBCA.saldoAkhir)}</b></div>
</Card>;})}
</div>}

{/* ── MODAL EDIT SETORAN ── */}
{editSetor&&!showSlip&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:500,padding:20,border:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.wht,marginBottom:14,fontSize:14}}>✏️ Edit Setoran — {fDs(editSetor.tanggal)}</div>
{(()=>{
var ef=editSetorF||editSetor;
var setEf=setEditSetorF;
return <div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
<Inp label="Tanggal Setor" type="date" value={ef.tanggal||""} onChange={v=>setEf({...ef,tanggal:v})}/>
<Sel label="Bank" value={ef.bank||"BSI"} onChange={v=>setEf({...ef,bank:v})} opts={[{v:"BSI",l:"BSI"},{v:"BCA",l:"BCA"}]}/>
<Inp label="Nominal (Rp)" type="number" value={String(ef.nominal||"")} onChange={v=>setEf({...ef,nominal:Number(v)})}/>
<Inp label="Penyetor" value={ef.penyetor||""} onChange={v=>setEf({...ef,penyetor:v})}/>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Pecahan Real (Tabel 2):</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:12}}>
{DENOMS.map(d=><div key={d} style={{display:"flex",alignItems:"center",gap:6,background:C.nav,borderRadius:6,padding:"4px 8px",border:"1px solid "+C.bdr}}>
<span style={{fontSize:11,color:C.gl2,minWidth:70}}>{fR(d)}</span>
<input type="number" value={(ef.pecahReal||ef.pecah||{})[d]||""} placeholder="0" onChange={e=>{var pr={...(ef.pecahReal||ef.pecah||{})};pr[d]=Number(e.target.value)||0;setEf({...ef,pecahReal:pr});}} style={{background:"transparent",border:"none",color:C.wht,fontSize:11,outline:"none",width:60,textAlign:"center"}}/>
</div>)}
</div>
<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
<button onClick={()=>{setEditSetor(null);setEditSetorF(null);}} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",color:C.gl2,cursor:"pointer",fontWeight:700,fontSize:12}}>Batal</button>
<button onClick={()=>{
var updated=(data.setoranBank||[]).map(x=>x.id===editSetor.id?{...editSetor,...ef,id:editSetor.id}:x);
setData(d=>({...d,setoranBank:updated}));
setEditSetor(null);setEditSetorF(null);
toast("✓ Setoran diperbarui!");
}} style={{background:C.glt,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>💾 Simpan</button>
<button onClick={()=>setShowSlip(true)} style={{background:C.blt,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>🖨️ Cetak Slip</button>
</div>
</div>;
})()}
</div>
</div>}

{/* ── MODAL SLIP SETORAN ── */}
{showSlip&&(editSetor||(data.setoranBank||[]).length>0)&&(()=>{
var last=editSetor||(data.setoranBank||[])[0];
var tglSetor=new Date(last.tanggal+"T00:00:00");
var hariLabel=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][tglSetor.getDay()];
var tglLabel=tglSetor.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
var noRek=last.bank==="BSI"?"812 69 2121 8":"—";
var logoSrc="data:image/png;base64,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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//Z";
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:"white",borderRadius:12,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,.5)",overflow:"hidden"}}>
<div id="_slip_setor" style={{background:"white",padding:20,fontFamily:"Arial,sans-serif",color:"#111",borderRadius:12}}>
{/* Header — format invoice */}
<div style={{background:"linear-gradient(135deg,#0a1f44 0%,#122d5e 100%)",padding:"14px 18px 12px",display:"flex",justifyContent:"space-between",alignItems:"center",margin:-20,marginBottom:0}}>
<div style={{display:"flex",alignItems:"center",gap:10}}>
{logoSrc?<img src={logoSrc} style={{height:44,objectFit:"contain"}} alt="Logo"/>:
<div style={{display:"flex",alignItems:"center",gap:10}}>
<div style={{display:"flex",gap:2,alignItems:"flex-end"}}>
<div style={{width:7,height:22,background:"#e53935",borderRadius:2}}/>
<div style={{width:7,height:16,background:"#1e88e5",borderRadius:2}}/>
<div style={{width:7,height:22,background:"#6ab04c",borderRadius:2}}/>
</div>
<div>
<div style={{fontSize:14,fontWeight:800,color:"white",letterSpacing:.3,lineHeight:1.1}}>{data.company?.nama||"PT. HOE TRANG SA"}</div>
<div style={{fontSize:8,fontWeight:600,color:"#e53935",letterSpacing:1.5,textTransform:"uppercase",marginTop:2}}>{data.company?.slogan||"DEALER LPG PERTAMINA"}</div>
</div>
</div>}
</div>
<div style={{display:"flex",alignItems:"center",gap:5}}>
{data.company?.logoPertamina?<img src={data.company.logoPertamina} style={{height:30,objectFit:"contain"}} alt="Pertamina"/>:
<div style={{display:"flex",alignItems:"center",gap:5}}>
<svg width={38} height={28} viewBox="0 0 60 42" fill="none">
<polygon points="18,0 32,14 18,28 4,14" fill="#E53935"/>
<polygon points="14,28 28,14 14,0 0,14" fill="#1565C0" opacity="0.85"/>
<polygon points="22,28 36,14 22,0 8,14" fill="#6AB04C" opacity="0.9"/>
</svg>
<span style={{fontSize:11,fontWeight:800,color:"white",letterSpacing:.4}}>PERTAMINA</span>
</div>}
</div>
</div>
{/* 3-color divider */}
<div style={{height:3,display:"flex",margin:"0 -20px 14px"}}>
<div style={{flex:1,background:"#1565c0"}}/><div style={{flex:1,background:"#6ab04c"}}/><div style={{flex:1,background:"#e53935"}}/>
</div>
<div style={{textAlign:"center",marginBottom:12}}>
<div style={{fontSize:13,fontWeight:700,color:"#0a1f44",letterSpacing:.5}}>RINCIAN SETORAN</div>
<div style={{fontSize:12,color:"#555"}}>{hariLabel} &nbsp;|&nbsp; {tglLabel}</div>
</div>
{/* Tabel pecahan */}
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:12}}>
<thead>
<tr style={{background:"#0a1f44"}}>
<th style={{padding:"6px 8px",color:"white",textAlign:"left",border:"1px solid #ccc"}}>PECAHAN</th>
<th style={{padding:"6px 8px",color:"white",textAlign:"center",border:"1px solid #ccc"}}>LBR</th>
<th style={{padding:"6px 8px",color:"white",textAlign:"right",border:"1px solid #ccc"}}>JUMLAH</th>
</tr>
</thead>
<tbody>
{DENOMS.map(d=>{
var lbr=Number(last.pecahReal?.[d]||last.pecah?.[d]||0);
var jml=lbr*d;
return <tr key={d} style={{borderBottom:"1px solid #e5e7eb"}}>
<td style={{padding:"5px 8px",color:"#111",border:"1px solid #ddd"}}>Rp {(d/1000).toFixed(0)}.000</td>
<td style={{padding:"5px 8px",textAlign:"center",color:lbr>0?"#0a1f44":"#aaa",fontWeight:lbr>0?700:400,border:"1px solid #ddd"}}>{lbr||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:jml>0?"#0a1f44":"#aaa",fontWeight:jml>0?700:400,border:"1px solid #ddd"}}>{jml>0?"Rp "+fR(jml).replace("Rp ",""):"Rp —"}</td>
</tr>;})}
<tr style={{background:"#0a1f44",fontWeight:700}}>
<td colSpan={2} style={{padding:"7px 8px",color:"white",border:"1px solid #ccc"}}>TOTAL</td>
<td style={{padding:"7px 8px",textAlign:"right",color:"white",fontSize:14,border:"1px solid #ccc"}}>Rp {fR(last.nominal).replace("Rp ","")}</td>
</tr>
</tbody>
</table>
{/* Footer info */}
<div style={{background:"#F8FAFC",borderRadius:6,padding:"10px 12px",border:"1px solid #E2E8F0",marginBottom:12}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
<span style={{fontSize:11,color:"#64748B",fontWeight:700}}>SETORAN LPG</span>
<span style={{fontSize:13,fontWeight:900,color:"#0a1f44"}}>{last.bank}</span>
</div>
{[["Penyetor",last.penyetor||penyetor],["Nama Nasabah","PT. HOE TRANG SA"],["Nomor Rekening",noRek]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
<span style={{fontSize:11,color:"#64748B"}}>{x[0]}</span>
<span style={{fontSize:11,fontWeight:700,color:"#0a1f44"}}>{x[1]}</span>
</div>)}
</div>
<div style={{textAlign:"center",fontSize:9,color:"#94A3B8"}}>
Jl. Jendral Sudirman No. 80 · Banda Aceh &nbsp;|&nbsp; 081269002121 / (0651) 21221
</div>
</div>
{/* Buttons */}
<div style={{padding:"10px 16px",display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center"}}>
<button onClick={()=>doPrint("_slip_setor")} style={{background:"#0a1f44",color:"white",border:"none",padding:"8px 16px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_slip_setor","Slip-Setor-"+last.bank+"-"+last.tanggal+".png")} style={{background:"#1D6A96",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_slip_setor")} style={{background:"#145A32",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>📋 Copy PNG</button>
<button onClick={()=>setShowSlip(false)} style={{background:"#6B7280",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>
</div>;
})()}

</div>;
}


var ALL_TABS=[
{id:"dashboard",label:"Dashboard",icon:"🏠"},{id:"penjualan",label:"Penjualan",icon:"🧾"},
{id:"piutang",label:"Piutang",icon:"💳"},{id:"setoran",label:"Setoran",icon:"💰"},
{id:"laporan",label:"Laporan",icon:"📊"},{id:"stok",label:"Stok",icon:"📦"},
{id:"pengeluaran",label:"Pengeluaran",icon:"💸"},{id:"tutupbuku",label:"Tutup Buku",icon:"📒"},
{id:"pelanggan",label:"Pelanggan",icon:"👥"},{id:"karyawan",label:"Karyawan",icon:"👤"},
{id:"absensi",label:"Absensi",icon:"📅"},{id:"payroll",label:"Payroll",icon:"💼"},
{id:"kas",label:"Kas & Bank",icon:"🏦"},{id:"do",label:"DO",icon:"🚚"},{id:"settings",label:"Pengaturan",icon:"⚙️"},
];
function getVisibleTabs(user){if(!user)return[];var rt=ROLE_TABS[user.role];if(rt===null||rt===undefined)return ALL_TABS;return ALL_TABS.filter(t=>rt.includes(t.id));}

export default function App(){
var[theme,setTheme]=useState(()=>{try{var s=localStorage.getItem("hts_theme");return s||"light";}catch(e){return"light";}});
var C=THEMES[theme]||THEMES.light;
useEffect(()=>{try{localStorage.setItem("hts_theme",theme);}catch(e){}},[theme]);
var[data,setData]=useState(()=>{
try{var s=localStorage.getItem("lpg_mgmt_v4");
if(s){var d=JSON.parse(s);

if(!d.stokKosong||typeof d.stokKosong!=="object")d.stokKosong={"5.5 kg":0,"12 kg":0,"50 kg":0};
["5.5 kg","12 kg","50 kg"].forEach(function(s){if(typeof d.stokKosong[s]==="undefined")d.stokKosong[s]=0;});
if(!d.totalTabung)d.totalTabung={"5.5 kg":0,"12 kg":0,"50 kg":0};
// Recalc totalTabung = isi + kosong + titip (logika benar)
(function(){var SIZES_M=["5.5 kg","12 kg","50 kg"];SIZES_M.forEach(function(s){var isi=(d.stock||{})[s]||0;var kosong=(d.stokKosong||{})[s]||0;var titip=0;(d.titipList||[]).forEach(function(t){var items=t.items&&t.items.length>0?t.items:(t.ukuran===s?[{qty:t.qty}]:[]);var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;items.forEach(function(it){if(!t.items||t.items.length===0?true:it.ukuran===s)titip+=m*Number(it.qty||0);});});d.totalTabung[s]=isi+kosong+Math.max(0,titip);});})();
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
if(t==="kas")return <KasBankMod {...props}/>;
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
console.log("[HTS-APP] Build: v4.3-fix-stok-titip — DO/Penj reverse, TitipTab fixed");
