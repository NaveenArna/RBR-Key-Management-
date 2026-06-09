import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import {
  getFirestore, doc, collection, getDocs, setDoc, updateDoc,
  deleteDoc, onSnapshot, writeBatch, serverTimestamp, query, orderBy, limit, where
} from "firebase/firestore";

// ─────────────────────────────────────────────────────────────────────────────
// Firebase Config
// ─────────────────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyCLwHkEWbqwlW_P01Q2j3Ply6p1QV-_sT4",
  authDomain: "rbr-key-management.firebaseapp.com",
  projectId: "rbr-key-management",
  storageBucket: "rbr-key-management.firebasestorage.app",
  messagingSenderId: "759345923480",
  appId: "1:759345923480:web:4b34d8b261a61d32d9d7ed"
};

const firebaseApp = initializeApp(firebaseConfig);
const db = getFirestore(firebaseApp);

// ─────────────────────────────────────────────────────────────────────────────
// Constants & Types
// ─────────────────────────────────────────────────────────────────────────────
const KEY_TYPES = ["Main Door Key","Mail Box Key","Pen drive","Smart Key","Other Key"] as const;
type KeyType = typeof KEY_TYPES[number];
const KEY_ICONS: Record<string,string> = {
  "Main Door Key":"🗝","Mail Box Key":"📬","Pen drive":"💾","Smart Key":"📱","Other Key":"🔑"
};
const DEFAULT_MAX = 200;
const DEFAULT_PASSCODE = "1234";
const STORAGE_KEY = "rbr_passcode";
const AUTO_LOCK_MS = 5 * 60 * 1000;

interface PropertyRow {
  "Lock Box No.": number;
  "Property Address": string;
  "Main Door Key": number;
  "Mail Box Key": number;
  "Pen drive": number;
  "Smart Key": number;
  "Other Key": number;
  _id: string;
}
interface BoxSettings { maxProps: number; }
interface BoxStat { props: number; keys: number; maxProps: number; pct: number; }
interface TxLog { id: string; ts: string; type: string; address: string; box: string; keyType: string; qty: number; user?: string; }
interface AddrEntry { addr: string; box: string; }

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────
function totalKeys(r: PropertyRow): number { return KEY_TYPES.reduce((s,k)=>s+(r[k]||0),0); }
function nowStr(): string { return new Date().toLocaleString("en-IN",{dateStyle:"short",timeStyle:"medium"}); }
function pctColor(p: number): string { return p>=100?"#e06060":p>=70?"#c8960c":"#50c880"; }
function getStoredPasscode(): string { try{return localStorage.getItem(STORAGE_KEY)||DEFAULT_PASSCODE;}catch{return DEFAULT_PASSCODE;} }
function savePasscode(p: string) { try{localStorage.setItem(STORAGE_KEY,p);}catch{} }

function levenshtein(a:string,b:string):number{
  const m=a.length,n=b.length;
  const dp:number[][]=Array.from({length:m+1},(_,i)=>Array.from({length:n+1},(_,j)=>i===0?j:j===0?i:0));
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++)
    dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1]);
  return dp[m][n];
}
function similarity(a:string,b:string):number{
  const al=a.toLowerCase(),bl=b.toLowerCase();
  if(bl.includes(al)||al.includes(bl)) return 1;
  const at=al.split(/[\s\-,]+/).filter(Boolean),bt=bl.split(/[\s\-,]+/).filter(Boolean);
  const hits=at.filter(t=>bt.some(b2=>b2.includes(t)||t.includes(b2)));
  const tok=at.length>0?hits.length/at.length:0;
  const ml=Math.max(al.length,bl.length);
  return tok*0.7+(ml>0?1-levenshtein(al,bl)/ml:1)*0.3;
}
function fuzzySearch(query:string,entries:AddrEntry[],limit=8):(AddrEntry&{score:number})[]{
  if(!query.trim()) return [];
  return entries.map(e=>({...e,score:similarity(query,e.addr)}))
    .filter(x=>x.score>0.12).sort((a,b)=>b.score-a.score).slice(0,limit);
}
function getBoxStats(data:Record<string,PropertyRow[]>,settings:Record<string,BoxSettings>):Record<string,BoxStat>{
  const r:Record<string,BoxStat>={};
  for(const [box,rows] of Object.entries(data)){
    const keys=rows.reduce((a,row)=>a+totalKeys(row),0);
    const maxProps=settings[box]?.maxProps||DEFAULT_MAX;
    r[box]={props:rows.length,keys,maxProps,pct:maxProps>0?Math.min(100,Math.round(rows.length/maxProps*100)):0};
  }
  return r;
}

// ─────────────────────────────────────────────────────────────────────────────
// Firebase helpers
// ─────────────────────────────────────────────────────────────────────────────
function safeId(id:string){return id.replace(/[^a-zA-Z0-9_-]/g,"_");}
async function fbSaveProperty(box:string, row:PropertyRow){
  // Use _id as doc path (it equals the Firestore doc ID from onSnapshot)
  await setDoc(doc(db,"boxes",safeId(box),"properties",row._id),{
    lockNo: row["Lock Box No."],
    address: row["Property Address"],
    mainDoor: row["Main Door Key"],
    mailBox: row["Mail Box Key"],
    penDrive: row["Pen drive"],
    smartKey: row["Smart Key"],
    otherKey: row["Other Key"],
    _id: row._id,
    updatedAt: serverTimestamp()
  });
}
async function fbDeleteProperty(box:string, id:string, address:string){
  // Strategy: find by BOTH document id AND by querying address
  // This handles any ID mismatch from old uploads
  try {
    // First try direct delete by id
    await deleteDoc(doc(db,"boxes",safeId(box),"properties",id));
  } catch(e) {
    console.warn("Direct delete failed, trying query by address:", e);
  }
  // Also query and delete any document with matching address (catches ID mismatches)
  const q = query(collection(db,"boxes",box,"properties"), where("address","==",address));
  const snapQ = await getDocs(q);
  for(const d of snapQ.docs){
    await deleteDoc(d.ref);
    console.log("Deleted by address query:", d.id);
  }
}
async function fbSaveBoxSettings(box:string, maxProps:number){
  await setDoc(doc(db,"settings",safeId(box)),{maxProps, updatedAt:serverTimestamp()},{merge:true});
}
async function fbSaveLog(log:TxLog){
  await setDoc(doc(db,"history",log.id),{...log, createdAt:serverTimestamp()});
}
async function fbClearBox(box:string){
  const snap=await getDocs(collection(db,"boxes",safeId(box),"properties"));
  const batch=writeBatch(db);
  snap.docs.forEach(d=>batch.delete(d.ref));
  await batch.commit();
}
async function fbClearHistory(){
  const snap=await getDocs(collection(db,"history"));
  const batch=writeBatch(db);
  snap.docs.forEach(d=>batch.delete(d.ref));
  await batch.commit();
}

function fbRowToProperty(d:any, docId?:string): PropertyRow {
  // ALWAYS use the Firestore document ID as _id
  // This guarantees delete/update always finds the right document
  return {
    "Lock Box No.": d.lockNo||0,
    "Property Address": d.address||"",
    "Main Door Key": d.mainDoor||0,
    "Mail Box Key": d.mailBox||0,
    "Pen drive": d.penDrive||0,
    "Smart Key": d.smartKey||0,
    "Other Key": d.otherKey||0,
    _id: docId||d._id||d.id  // docId = actual Firestore doc ID
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Initial seed data (uploaded once to Firebase)
// ─────────────────────────────────────────────────────────────────────────────
const SEED_DATA: Record<string,any[]> = {"BOX 1":[{"Lock Box No.":1,"Property Address":"1628 Aspire St-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-0"},{"Lock Box No.":2,"Property Address":"2418 Arbor Loop Dr-28217","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-1"},{"Lock Box No.":3,"Property Address":"2837 Ensemble Court","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-2"},{"Lock Box No.":4,"Property Address":"113 Crossvine Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-3"},{"Lock Box No.":5,"Property Address":"2783 Berkhamstead Circle","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-4"},{"Lock Box No.":6,"Property Address":"9128 Lowfalls Lane","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-5"},{"Lock Box No.":7,"Property Address":"7812 Nelson Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-6"},{"Lock Box No.":8,"Property Address":"772 Earhart St NW","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-7"},{"Lock Box No.":9,"Property Address":"2877 Yaeger Drive NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-8"},{"Lock Box No.":10,"Property Address":"242 Abersham Drive-28115","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-9"},{"Lock Box No.":11,"Property Address":"5150 Hyrule Dr-28262","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-10"},{"Lock Box No.":12,"Property Address":"12519 Bryton ridge Pkwy","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-11"},{"Lock Box No.":13,"Property Address":"12115 Devon Square Court-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-12"},{"Lock Box No.":14,"Property Address":"10159 Chatham Run Lane","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-13"},{"Lock Box No.":15,"Property Address":"11027 Pagebrook Ln-28214","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-14"},{"Lock Box No.":16,"Property Address":"1572 Forkhorn Dr-28110","Main Door Key":0,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-15"},{"Lock Box No.":17,"Property Address":"1227 Southern Sugar Dr-28262","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-16"},{"Lock Box No.":18,"Property Address":"2026 Tears Ln-28217","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-17"},{"Lock Box No.":19,"Property Address":"10310 Glenmere Creek Circle-28262","Main Door Key":0,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-18"},{"Lock Box No.":20,"Property Address":"14024 Singleleaf Lane-28278","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-19"},{"Lock Box No.":21,"Property Address":"14029 Whistling Tear Dr-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-20"},{"Lock Box No.":22,"Property Address":"1671 Spears Drive-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-21"},{"Lock Box No.":23,"Property Address":"11255 Bryton Parkway-28078","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-22"},{"Lock Box No.":24,"Property Address":"10212 University Park Ln","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-23"},{"Lock Box No.":25,"Property Address":"2301 Endeavor Rn","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-24"},{"Lock Box No.":26,"Property Address":"1621 Nia Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-25"},{"Lock Box No.":27,"Property Address":"11229 Smokethron Dr","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-26"},{"Lock Box No.":28,"Property Address":"2158 Holden Avenue Southwest-28025 ","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-27"},{"Lock Box No.":29,"Property Address":"1312 Killashee Ct-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-28"},{"Lock Box No.":30,"Property Address":"3667 Ascott Commons Ln","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-29"},{"Lock Box No.":31,"Property Address":"12562 Mcgrath Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-30"},{"Lock Box No.":32,"Property Address":"1706 Spears Dr NW-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-31"},{"Lock Box No.":33,"Property Address":"2765 Yeager Drive NW ","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-32"},{"Lock Box No.":34,"Property Address":"2853 Yeager Drive NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-33"},{"Lock Box No.":35,"Property Address":"1729 Braemar Village Dr-Monroe","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-34"},{"Lock Box No.":36,"Property Address":"10154 Chatham Run Ln","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-35"},{"Lock Box No.":37,"Property Address":"2033 old Rivers Rd-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-36"},{"Lock Box No.":38,"Property Address":"7879 Iron Road(Sherrills Road)","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-37"},{"Lock Box No.":39,"Property Address":"14305 Piper Landing Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-38"},{"Lock Box No.":40,"Property Address":"8006 Ramsburg Dr","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-39"},{"Lock Box No.":41,"Property Address":"1059 Grays Mill Road","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-40"},{"Lock Box No.":42,"Property Address":"1520 Tranquility Ave NW","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-41"},{"Lock Box No.":43,"Property Address":"8924 Connoer Hall Ave","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-42"},{"Lock Box No.":44,"Property Address":"1815 Arbor Vista Dr-28262","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-43"},{"Lock Box No.":45,"Property Address":"4133 Lurelin Lane","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-44"},{"Lock Box No.":46,"Property Address":"10431 Bunclody Dr-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-45"},{"Lock Box No.":47,"Property Address":"7714 Nelson Rd","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-46"},{"Lock Box No.":48,"Property Address":"9294 Perseverance Dr-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-47"},{"Lock Box No.":49,"Property Address":"15126 Windy Plains Rd-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-48"},{"Lock Box No.":50,"Property Address":"1399 Cedardale Ln-28037","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-49"},{"Lock Box No.":51,"Property Address":"1474 Olive Hill Ave NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-50"},{"Lock Box No.":52,"Property Address":"274 Halton Crossing Dr SW","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-51"},{"Lock Box No.":53,"Property Address":"4120 County DownAvenue-28081","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-52"},{"Lock Box No.":54,"Property Address":"2136 Laurens Dr-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-53"},{"Lock Box No.":55,"Property Address":"2301 Endeavor Run-28269","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-54"},{"Lock Box No.":56,"Property Address":"1765 Evergreen Drive-28208","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-55"},{"Lock Box No.":57,"Property Address":"3683 Backwater St-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-56"},{"Lock Box No.":58,"Property Address":"4460 Sourwood Court","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-57"},{"Lock Box No.":59,"Property Address":"5720 Rivulet Way","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-58"},{"Lock Box No.":60,"Property Address":"5543 Worrell Way-Kannapolis","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-59"},{"Lock Box No.":61,"Property Address":"7828 Nelson Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-60"},{"Lock Box No.":62,"Property Address":"104 Winterberry St-28117","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-61"},{"Lock Box No.":63,"Property Address":"7106 Waterwheel St SW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-62"},{"Lock Box No.":64,"Property Address":"9646 Munsing Drive","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-63"},{"Lock Box No.":65,"Property Address":"2128 Blue Sky Mdws Dr-28110","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-64"},{"Lock Box No.":66,"Property Address":"1225 Colgher St-28227","Main Door Key":1,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-65"},{"Lock Box No.":67,"Property Address":"8377 Breton Way-28075","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-66"},{"Lock Box No.":68,"Property Address":"1422 Newell Towns Ln-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-67"},{"Lock Box No.":69,"Property Address":"12022 Elizabeth Madison Court-28277","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-68"},{"Lock Box No.":70,"Property Address":"8226 Merryvale Ln-28214","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-69"},{"Lock Box No.":71,"Property Address":"9006 Clayton Alley-28027","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-70"},{"Lock Box No.":72,"Property Address":"5142 Elementor View Dr-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-71"},{"Lock Box No.":73,"Property Address":"9621 Cherry Meadow Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-72"},{"Lock Box No.":74,"Property Address":"2019 Sage Park Dr-28217","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-73"},{"Lock Box No.":75,"Property Address":"5847 Strathmore Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-74"},{"Lock Box No.":76,"Property Address":"5430 Kyndall Walk way-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-75"},{"Lock Box No.":77,"Property Address":"2220 Belterra Dr-28216","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-76"},{"Lock Box No.":78,"Property Address":"784 Earhart St NW","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-77"},{"Lock Box No.":79,"Property Address":"6410 mallard View Ln-28269","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-78"},{"Lock Box No.":80,"Property Address":"13034 Garren Vw Ln-28278","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-79"},{"Lock Box No.":81,"Property Address":"2207 Restina Drive-28173","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-80"},{"Lock Box No.":82,"Property Address":"9308 Widden Way NC-28269","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-81"},{"Lock Box No.":83,"Property Address":"9161 Harwen Ln","Main Door Key":3,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-82"},{"Lock Box No.":84,"Property Address":"9215 Harwen Ln","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-83"},{"Lock Box No.":85,"Property Address":"8608 Lavender PI","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-84"},{"Lock Box No.":86,"Property Address":"619 Breckenridge Rd-Kannapolis","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-85"},{"Lock Box No.":87,"Property Address":"4370 Evening Trail-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-86"},{"Lock Box No.":88,"Property Address":"6405 Prosperity Church Road-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-87"},{"Lock Box No.":89,"Property Address":"9225 Delancey Ln SW","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-88"},{"Lock Box No.":90,"Property Address":"790 Earhart St NW","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-89"},{"Lock Box No.":91,"Property Address":"3603 Backwater Street","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-90"},{"Lock Box No.":92,"Property Address":"5243 Brailey Cir-Kannapolis","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-91"},{"Lock Box No.":93,"Property Address":"5601 Stafford Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-92"},{"Lock Box No.":94,"Property Address":"5005 Jamie Sloop Ln","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-93"},{"Lock Box No.":95,"Property Address":"4060 Backwater Street","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-94"},{"Lock Box No.":96,"Property Address":"15117 Windy Plains Road","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-95"},{"Lock Box No.":97,"Property Address":"439 Sweet Shrub Court-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-96"},{"Lock Box No.":98,"Property Address":"5043 Sunnybrae PI-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-97"},{"Lock Box No.":99,"Property Address":"3125 Hutton Gardens-28269","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-98"},{"Lock Box No.":100,"Property Address":"3321 Finch Borough Ct-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-99"},{"Lock Box No.":101,"Property Address":"2715 Yeager Drive NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-100"},{"Lock Box No.":102,"Property Address":"4860 Pepper Dr-Harriburg","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-101"},{"Lock Box No.":103,"Property Address":"4053 Lawnview Dr","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-102"},{"Lock Box No.":104,"Property Address":"9008 Clayton Aly-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-103"},{"Lock Box No.":105,"Property Address":"2627 Silverthorn Dr-28273","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-104"},{"Lock Box No.":106,"Property Address":"4015 Meadow Green Dr","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-105"},{"Lock Box No.":107,"Property Address":"4051 Zilker park D-28217","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-106"},{"Lock Box No.":108,"Property Address":"5864 Coulee Ln","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-107"},{"Lock Box No.":109,"Property Address":"2011 Pippen Ave","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-108"},{"Lock Box No.":110,"Property Address":"5044 Falstone Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-109"},{"Lock Box No.":111,"Property Address":"3650 Backwater St","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-110"},{"Lock Box No.":112,"Property Address":"5428 Kinbridge Dr","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-111"},{"Lock Box No.":113,"Property Address":"5629 Stafford Road","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-112"},{"Lock Box No.":114,"Property Address":"5017 Sovereignty Ct","Main Door Key":0,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-113"},{"Lock Box No.":115,"Property Address":"5029 Falstone Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-114"},{"Lock Box No.":116,"Property Address":"3670 Backwater St","Main Door Key":6,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-115"},{"Lock Box No.":117,"Property Address":"4087 Long Arrow Dr-28025","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-116"},{"Lock Box No.":118,"Property Address":"5596 Stafford Rd","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-117"},{"Lock Box No.":119,"Property Address":"2136 Highland Knoll Drive","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-118"},{"Lock Box No.":120,"Property Address":"7836 Nelson Road-Mint Hill","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-119"},{"Lock Box No.":121,"Property Address":"736 Lock Harven Dr NW","Main Door Key":3,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-120"},{"Lock Box No.":122,"Property Address":"6122 Faron Way","Main Door Key":5,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":1,"_id":"BOX 1-121"},{"Lock Box No.":123,"Property Address":"12111 Devon Square Court","Main Door Key":4,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-122"},{"Lock Box No.":124,"Property Address":"2183 Falling Acorn  Ln-28027","Main Door Key":3,"Mail Box Key":1,"Pen drive":1,"Smart Key":0,"Other Key":0,"_id":"BOX 1-123"},{"Lock Box No.":125,"Property Address":"5253 Brailey Cir","Main Door Key":1,"Mail Box Key":2,"Pen drive":1,"Smart Key":1,"Other Key":0,"_id":"BOX 1-124"},{"Lock Box No.":126,"Property Address":"6105 Starview Terrace-28216","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-125"},{"Lock Box No.":127,"Property Address":"2504 Abundance Ln-28173","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-126"},{"Lock Box No.":128,"Property Address":"5763 coulee ln","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-127"},{"Lock Box No.":129,"Property Address":"9004 Clayton Aly-Concord-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-128"},{"Lock Box No.":130,"Property Address":"7145 Bentz St-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-129"},{"Lock Box No.":131,"Property Address":"107 Jameson Pk Dr-28116","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-130"},{"Lock Box No.":132,"Property Address":"13131 Hampton Bay Ln-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-131"},{"Lock Box No.":133,"Property Address":"3011 Glenn Hope Wy-28104","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-132"},{"Lock Box No.":134,"Property Address":"7852 Nelson Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-133"},{"Lock Box No.":135,"Property Address":"10732 Alvarado Way28277","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-134"},{"Lock Box No.":136,"Property Address":"8158 Rudalph Road-28216","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-135"},{"Lock Box No.":137,"Property Address":"6519 Revolutionary Trail-28217","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-136"},{"Lock Box No.":138,"Property Address":"2778 Yeager Dr NW-28025","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-137"},{"Lock Box No.":139,"Property Address":"239 Harpers Run Ln-28104","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-138"},{"Lock Box No.":140,"Property Address":"9146 Redmond Trace Rd-28277","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-139"},{"Lock Box No.":141,"Property Address":"10320 Ebbets Rd-28273","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-140"},{"Lock Box No.":142,"Property Address":"1604 Simril Ct-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-141"},{"Lock Box No.":143,"Property Address":"14418 Target Ln-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-142"},{"Lock Box No.":144,"Property Address":"2871 Yeager Dr NW-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-143"},{"Lock Box No.":145,"Property Address":"3309 Linetender Dr-28036","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-144"},{"Lock Box No.":146,"Property Address":"6104 Balham Ct-28215","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-145"},{"Lock Box No.":147,"Property Address":"1508 Blanche St-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-146"},{"Lock Box No.":148,"Property Address":"1272 Scarlet Firethrone Ave NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-147"},{"Lock Box No.":149,"Property Address":"18016 Wilbanks Dr-28278","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-148"},{"Lock Box No.":150,"Property Address":"7022 Walnut Branch Ln-28277","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-149"},{"Lock Box No.":151,"Property Address":"1706 Blanche St-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-150"},{"Lock Box No.":152,"Property Address":"3725 Burntwood Ct-28227","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-151"},{"Lock Box No.":153,"Property Address":"4225 Stoneygreen Ln-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-152"},{"Lock Box No.":154,"Property Address":"7067 Waterwheel St SW-28025","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-153"},{"Lock Box No.":155,"Property Address":"8113 Murray Br Dr-28216","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-154"},{"Lock Box No.":156,"Property Address":"8403 Bristle Toe","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-155"},{"Lock Box No.":157,"Property Address":"10730 Tuff Ln-28036","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-156"},{"Lock Box No.":158,"Property Address":"13932 Castle Nook Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-157"},{"Lock Box No.":159,"Property Address":"1616 Wilburn Park Lane-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-158"},{"Lock Box No.":160,"Property Address":"3239 Hampton Bay Lane","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-159"},{"Lock Box No.":161,"Property Address":"12511 Bryton Ridge Pkwy-28078","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-160"},{"Lock Box No.":162,"Property Address":"7402 Dover Mill Dr SW-28025","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-161"},{"Lock Box No.":163,"Property Address":"2116 Clapham Ct-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-162"},{"Lock Box No.":164,"Property Address":"3885 Cullen Meadows Dr-28036","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-163"},{"Lock Box No.":165,"Property Address":"9028 Tamarron Drive-28277","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-164"},{"Lock Box No.":166,"Property Address":"6998 Founders Way-28075","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-165"},{"Lock Box No.":167,"Property Address":"4107 Black Ct-28075","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-166"},{"Lock Box No.":168,"Property Address":"11091 River Oaks Dr Nw","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-167"},{"Lock Box No.":169,"Property Address":"3823 Dahalia Drive","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-168"},{"Lock Box No.":170,"Property Address":"586 Brook Haven-Fort Mill","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-169"},{"Lock Box No.":171,"Property Address":"4110 County Down Avenue-28081","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-170"},{"Lock Box No.":172,"Property Address":"369 Abington St","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-171"},{"Lock Box No.":173,"Property Address":"11346 Cedar Walk lane","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-172"},{"Lock Box No.":174,"Property Address":"1778 Braemer Village Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-173"},{"Lock Box No.":175,"Property Address":"2306 Rachelwood Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-174"},{"Lock Box No.":176,"Property Address":"3530 Secrest Landing","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-175"},{"Lock Box No.":177,"Property Address":"3615 Secrest Landing","Main Door Key":4,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-176"},{"Lock Box No.":178,"Property Address":"2809 Ava Ave","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-177"},{"Lock Box No.":179,"Property Address":"4155 Stream Dale Cir NW","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-178"},{"Lock Box No.":180,"Property Address":"4509 morning Dew","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-179"},{"Lock Box No.":181,"Property Address":"10604 Haddington Drive","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-180"},{"Lock Box No.":182,"Property Address":"3576 Nimbell Rd-28110","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-181"},{"Lock Box No.":183,"Property Address":"2808 Twinberry Ln","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":1,"_id":"BOX 1-182"},{"Lock Box No.":184,"Property Address":"13319 Savaine St","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-183"},{"Lock Box No.":185,"Property Address":"2160 Blue Sky Meadows","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-184"},{"Lock Box No.":186,"Property Address":"10640 Simril Ct","Main Door Key":6,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-185"},{"Lock Box No.":187,"Property Address":"4514 Dover Nest Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-186"},{"Lock Box No.":188,"Property Address":"1705 Hawthorne Lane","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-187"},{"Lock Box No.":189,"Property Address":"412 Wild Dove Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 1-188"},{"Lock Box No.":190,"Property Address":"2624 Snap Dragon Dr","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":1,"_id":"BOX 1-189"}],"BOX 2":[{"Lock Box No.":1,"Property Address":"744 Lock Haven Drive NW-28028","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-0"},{"Lock Box No.":2,"Property Address":"7109 Pennyroyal Wy-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-1"},{"Lock Box No.":3,"Property Address":"3240 Stelfox St-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-2"},{"Lock Box No.":4,"Property Address":"4337 Smt Wds Dr-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-3"},{"Lock Box No.":5,"Property Address":"3719 Wave Rock Ct-29707","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-4"},{"Lock Box No.":6,"Property Address":"2241 Cobble Ct-28110","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-5"},{"Lock Box No.":7,"Property Address":"5605 Stafford Rd-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-6"},{"Lock Box No.":8,"Property Address":"7011 Waterwheel St SW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-7"},{"Lock Box No.":9,"Property Address":"3483 Backwater St-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-8"},{"Lock Box No.":10,"Property Address":"2321 Creekmere Lane-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-9"},{"Lock Box No.":11,"Property Address":"1453 Harleston St-28079","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-10"},{"Lock Box No.":12,"Property Address":"1726 Braemar Village Dr-28110","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-11"},{"Lock Box No.":13,"Property Address":"1728 Aspire St-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-12"},{"Lock Box No.":14,"Property Address":"9008 Catboat St-28078","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-13"},{"Lock Box No.":15,"Property Address":"2751 Bramble Ridge Ct-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-14"},{"Lock Box No.":16,"Property Address":"14032 Penbury Ln-28278","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-15"},{"Lock Box No.":17,"Property Address":"3526 Cramer Crk Dr-28056","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-16"},{"Lock Box No.":18,"Property Address":"9338 Mallard Mills Dr-28262","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-17"},{"Lock Box No.":19,"Property Address":"2833 Statesville Ave-28206","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-18"},{"Lock Box No.":20,"Property Address":"3432 Blf Hill Ln-28215","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-19"},{"Lock Box No.":21,"Property Address":"10988 Flyreel PI-28036","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-20"},{"Lock Box No.":22,"Property Address":"18010 Stark Way-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-21"},{"Lock Box No.":23,"Property Address":"11126 Green Spring Drive-28078","Main Door Key":0,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-22"},{"Lock Box No.":24,"Property Address":"1713 Unison Dr-28262","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-23"},{"Lock Box No.":25,"Property Address":"10142 Black Locust Ln-28075","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-24"},{"Lock Box No.":26,"Property Address":"17113 Sand Bank Rd-28278","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":2,"_id":"BOX 2-25"},{"Lock Box No.":27,"Property Address":"8227 Cousins Ct-29707","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-26"},{"Lock Box No.":28,"Property Address":"19010 Direwolf Cove-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-27"},{"Lock Box No.":29,"Property Address":"2567 Cornelius PI NW-28027","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-28"},{"Lock Box No.":30,"Property Address":"1735 Braemar Village Dr","Main Door Key":3,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-29"},{"Lock Box No.":31,"Property Address":"2418 Arbor Loop Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-30"},{"Lock Box No.":32,"Property Address":"4004 Petersburg Dr-28173","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-31"},{"Lock Box No.":33,"Property Address":"7005 Ludell Ln-28215","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-32"},{"Lock Box No.":34,"Property Address":"8025 Bristle Toe Lane-28277","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-33"},{"Lock Box No.":35,"Property Address":"1572 Forkhorn Dr","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-34"},{"Lock Box No.":36,"Property Address":"5676 Clear Creek Ln-28215","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-35"},{"Lock Box No.":37,"Property Address":"5726 Rivulet Wy","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-36"},{"Lock Box No.":38,"Property Address":"117 Snead Rd-29715","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-37"},{"Lock Box No.":39,"Property Address":"7832 Nelson Road","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-38"},{"Lock Box No.":40,"Property Address":"113 Crossvine Dr-28117","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-39"},{"Lock Box No.":41,"Property Address":"14924 Nicolas Hall Dr","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-40"},{"Lock Box No.":42,"Property Address":"1011 Ketchum Ct","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-41"},{"Lock Box No.":43,"Property Address":"3053 Summerfield Ridge Ln-28105","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-42"},{"Lock Box No.":44,"Property Address":"234 Quinn Rd-Matthews","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-43"},{"Lock Box No.":45,"Property Address":"14054 Singleleaf Ln-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-44"},{"Lock Box No.":46,"Property Address":"12022 Elizabeth Madison Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-45"},{"Lock Box No.":47,"Property Address":"1716 Vanderlyn St-Monroe","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-46"},{"Lock Box No.":48,"Property Address":"10223 Kelso Ct","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":1,"_id":"BOX 2-47"},{"Lock Box No.":49,"Property Address":"9769 Oaklawn Blvd NW-28078","Main Door Key":1,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-48"},{"Lock Box No.":50,"Property Address":"2508 Mccurdy Trail-28269","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-49"},{"Lock Box No.":51,"Property Address":"9320 Alice McGinn Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-50"},{"Lock Box No.":52,"Property Address":"409 Wild Dove Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-51"},{"Lock Box No.":53,"Property Address":"3836 Memorial Park Way","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-52"},{"Lock Box No.":54,"Property Address":"239 Harpers Run Ln","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-53"},{"Lock Box No.":55,"Property Address":"416 Wild Dove Ct","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-54"},{"Lock Box No.":56,"Property Address":"11753 Mesquite Rd-28078","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-55"},{"Lock Box No.":57,"Property Address":"11622 Red Rust Lane","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-56"},{"Lock Box No.":58,"Property Address":"2802 Azalea Hills Dr","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-57"},{"Lock Box No.":59,"Property Address":"4126 Salient St-28205","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-58"},{"Lock Box No.":60,"Property Address":"1312 Secrest Commons Dr","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":1,"_id":"BOX 2-59"},{"Lock Box No.":61,"Property Address":"6319 Southgrove St-28277","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-60"},{"Lock Box No.":62,"Property Address":"4217 Lake Rd-28269","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-61"},{"Lock Box No.":63,"Property Address":"17112 Carolina Hickory Dr-28078","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":2,"_id":"BOX 2-62"},{"Lock Box No.":64,"Property Address":"3958 Rothwood Ln-28075","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-63"},{"Lock Box No.":65,"Property Address":"2809 Ava Ave","Main Door Key":7,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-64"},{"Lock Box No.":66,"Property Address":"3220 Lilac Grove Dr-28269","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-65"},{"Lock Box No.":67,"Property Address":"1725 Aspire Street-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-66"},{"Lock Box No.":68,"Property Address":"6258 tea Olive Dr-28075","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-67"},{"Lock Box No.":69,"Property Address":"8226 Merryvale Ln","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-68"},{"Lock Box No.":70,"Property Address":"7106 Waterwheel St SW-28025","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-69"},{"Lock Box No.":71,"Property Address":"2452 Royal York Ave","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-70"},{"Lock Box No.":72,"Property Address":"7848 Nelson Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-71"},{"Lock Box No.":73,"Property Address":"7928 Denmark Road CLT","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-72"},{"Lock Box No.":74,"Property Address":"7816 Nelson Rd","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-73"},{"Lock Box No.":75,"Property Address":"16341 Leading St","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-74"},{"Lock Box No.":76,"Property Address":"3234 Hiram St-28208","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-75"},{"Lock Box No.":77,"Property Address":"240 Harpers Run Ln-Matthews","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-76"},{"Lock Box No.":78,"Property Address":"133 Harpers Run Ln","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-77"},{"Lock Box No.":79,"Property Address":"10613 Bere Island","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-78"},{"Lock Box No.":80,"Property Address":"5504 Kins Bridge Dr,Mint Hill","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-79"},{"Lock Box No.":81,"Property Address":"9615 Weikert Rd-28215","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-80"},{"Lock Box No.":82,"Property Address":"10327 Snowbell Ct-28215","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-81"},{"Lock Box No.":83,"Property Address":"5110 Hyrule Dr-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-82"},{"Lock Box No.":84,"Property Address":"2849 Aubrey St-Monroe","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-83"},{"Lock Box No.":85,"Property Address":"1617 Blanche St-28262","Main Door Key":3,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-84"},{"Lock Box No.":86,"Property Address":"14917 Baldridge Dr-28078","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-85"},{"Lock Box No.":87,"Property Address":"3530 Secrest Landing","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-86"},{"Lock Box No.":88,"Property Address":"5069 Grain Orchard Rd-28079","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-87"},{"Lock Box No.":89,"Property Address":"2881 Yeager Dr NW-concord","Main Door Key":0,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-88"},{"Lock Box No.":90,"Property Address":"4051 Zilker Park Dr","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-89"},{"Lock Box No.":91,"Property Address":"2314 Donnelly Hills Ln-28262","Main Door Key":3,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-90"},{"Lock Box No.":92,"Property Address":"6329 Marquam PI-28215","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-91"},{"Lock Box No.":93,"Property Address":"3236 Stelfox St-28262","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-92"},{"Lock Box No.":94,"Property Address":"3160 Lilac Grove Dr-28269","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-93"},{"Lock Box No.":95,"Property Address":"2026 Tears Ln-28217","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-94"},{"Lock Box No.":96,"Property Address":"4135 Summit Woods Dr-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-95"},{"Lock Box No.":97,"Property Address":"15720 Country House Street","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-96"},{"Lock Box No.":98,"Property Address":"11255 Bryton Park Wy","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-97"},{"Lock Box No.":99,"Property Address":"4117 Summit Woods Dr-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-98"},{"Lock Box No.":100,"Property Address":"1742 Blanche St-28262","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-99"},{"Lock Box No.":101,"Property Address":"5017 Sovereignty Ct-28205","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-100"},{"Lock Box No.":102,"Property Address":"10159 Chatham Run Ln","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-101"},{"Lock Box No.":103,"Property Address":"14040 Lake Home Ln-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-102"},{"Lock Box No.":104,"Property Address":"6607 Wildbrook Dr-28269","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-103"},{"Lock Box No.":105,"Property Address":"5912 faron Way-28262","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-104"},{"Lock Box No.":106,"Property Address":"4464 Millennium Ave-28217","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-105"},{"Lock Box No.":107,"Property Address":"3530 Alister Ave SW-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-106"},{"Lock Box No.":108,"Property Address":"14506 Crociani Dr-28277","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-107"},{"Lock Box No.":109,"Property Address":"7651 Fenn Way-29707","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-108"},{"Lock Box No.":110,"Property Address":"5015 Westmead Ln-28262","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-109"},{"Lock Box No.":111,"Property Address":"1361 Rainier Dr-29708","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-110"},{"Lock Box No.":112,"Property Address":"4143 Summit Woods Dr-28216","Main Door Key":2,"Mail Box Key":2,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-111"},{"Lock Box No.":113,"Property Address":"2352 Belterra Dr-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-112"},{"Lock Box No.":114,"Property Address":"5710 Gulch PI-28215","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-113"},{"Lock Box No.":115,"Property Address":"5110 Friendly Baptist Church Rd-28079","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-114"},{"Lock Box No.":116,"Property Address":"4938 Stffordshire Ln-28213","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-115"},{"Lock Box No.":117,"Property Address":"2603 Bridle Brook Way-28270","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-116"},{"Lock Box No.":118,"Property Address":"15144 Jade St-28277","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-117"},{"Lock Box No.":119,"Property Address":"2222 Transatlantic Ave-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-118"},{"Lock Box No.":120,"Property Address":"7322 Copper Beech Trce-28273","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-119"},{"Lock Box No.":121,"Property Address":"1939 Stallings Rd-28104","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-120"},{"Lock Box No.":122,"Property Address":"6635 Wildbrook Dr-28078","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-121"},{"Lock Box No.":123,"Property Address":"7214 Waterwheel St SW-28025","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-122"},{"Lock Box No.":124,"Property Address":"5504 Joshua Cain Rd-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-123"},{"Lock Box No.":125,"Property Address":"2203 Autumn Olive Ln-28104","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-124"},{"Lock Box No.":126,"Property Address":"3419 Secrest Lndg-28110","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-125"},{"Lock Box No.":127,"Property Address":"8113 Murray Br Dr-28216","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-126"},{"Lock Box No.":128,"Property Address":"4118 Salient Street NC-28205","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-127"},{"Lock Box No.":129,"Property Address":"2885 Yeager Dr NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-128"},{"Lock Box No.":130,"Property Address":"10617 Bunclody Dr-28213","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-129"},{"Lock Box No.":131,"Property Address":"4037 Stoneygreen Ln-28215","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-130"},{"Lock Box No.":132,"Property Address":"7311 Mitzi Deborah Ln-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-131"},{"Lock Box No.":133,"Property Address":"4124 Salient St-28205","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-132"},{"Lock Box No.":134,"Property Address":"17322 Carolina Hickory Dr","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-133"},{"Lock Box No.":135,"Property Address":"9529 Teamwork St NW","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-134"},{"Lock Box No.":136,"Property Address":"493 Twelve Oaks Ln-29708","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-135"},{"Lock Box No.":137,"Property Address":"3483 Backwater St-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-136"},{"Lock Box No.":138,"Property Address":"5108 Carrick St-28213","Main Door Key":3,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-137"},{"Lock Box No.":139,"Property Address":"19016 Yellow Birch Dr-28278","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-138"},{"Lock Box No.":140,"Property Address":"205 Limerick Rd-28115 unit-D","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-139"},{"Lock Box No.":141,"Property Address":"4698 Selhurst Dr-SC-29707","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-140"},{"Lock Box No.":142,"Property Address":"2034 Highland Park Dr-28269","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-141"},{"Lock Box No.":143,"Property Address":"348 Cranford Dr-28134","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":2,"_id":"BOX 2-142"},{"Lock Box No.":144,"Property Address":"9046 Evercrisp Ln-28215","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-143"},{"Lock Box No.":145,"Property Address":"808 Gable Oak Ln #59-29708","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-144"},{"Lock Box No.":146,"Property Address":"490 Tayberry Ln-29715","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-145"},{"Lock Box No.":147,"Property Address":"4709 Morning Dew Ct-28269","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-146"},{"Lock Box No.":148,"Property Address":"6126 Russo Ct-29720","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-147"},{"Lock Box No.":149,"Property Address":"246 Ferebee PI-28213","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-148"},{"Lock Box No.":150,"Property Address":"11007 Discovery Dr NW-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-149"},{"Lock Box No.":151,"Property Address":"9626 Kenneth Glenn Dr-28213","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-150"},{"Lock Box No.":152,"Property Address":"1618 Swallow Tail Dr-28012","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-151"},{"Lock Box No.":153,"Property Address":"5984 River Meadow Ct-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-152"},{"Lock Box No.":154,"Property Address":"2835 Berkhamstead Cir-28027","Main Door Key":1,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-153"},{"Lock Box No.":155,"Property Address":"17034 Patron Drive -28273","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-154"},{"Lock Box No.":156,"Property Address":"8046 Saluda Dr-28269","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-155"},{"Lock Box No.":157,"Property Address":"727 Dynamo St NW-28027","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-156"},{"Lock Box No.":158,"Property Address":"8147 Paw Club Dr-28214","Main Door Key":1,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-157"},{"Lock Box No.":159,"Property Address":"7433 Sienna Heights PI-28213","Main Door Key":2,"Mail Box Key":1,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-158"},{"Lock Box No.":160,"Property Address":"3629 Edisto PI-Monroe","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-159"},{"Lock Box No.":161,"Property Address":"5803 Camp Ct-28025","Main Door Key":3,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-160"},{"Lock Box No.":162,"Property Address":"16849 Greenlawn Hills Ct-28213","Main Door Key":2,"Mail Box Key":0,"Pen drive":0,"Smart Key":0,"Other Key":0,"_id":"BOX 2-161"}]};

// ─────────────────────────────────────────────────────────────────────────────
// LockScreen component
// ─────────────────────────────────────────────────────────────────────────────
interface LockScreenProps { onUnlock:()=>void; mode?:"unlock"|"change"; onCancel?:()=>void; }
function LockScreen({onUnlock,mode="unlock",onCancel}:LockScreenProps){
  const [digits,setDigits]=useState<string[]>([]);
  const [step,setStep]=useState<"enter"|"new1"|"new2">(mode==="change"?"new1":"enter");
  const [newPin,setNewPin]=useState("");
  const [shake,setShake]=useState(false);
  const [message,setMessage]=useState(mode==="change"?"Enter new 4-digit PIN":"Enter your PIN");
  const [success,setSuccess]=useState(false);

  function triggerShake(){setShake(true);setTimeout(()=>{setShake(false);setDigits([]);},500);}
  function handleDigit(d:string){
    if(digits.length>=4) return;
    const next=[...digits,d]; setDigits(next);
    if(next.length===4) setTimeout(()=>evaluate(next.join("")),120);
  }
  function evaluate(pin:string){
    if(step==="enter"){
      if(pin===getStoredPasscode()){setSuccess(true);setTimeout(onUnlock,300);}
      else{setMessage("Incorrect PIN - try again");triggerShake();}
    } else if(step==="new1"){
      setNewPin(pin);setDigits([]);setStep("new2");setMessage("Confirm new PIN");
    } else if(step==="new2"){
      if(pin===newPin){savePasscode(pin);setSuccess(true);setMessage("✓ PIN updated!");setTimeout(onUnlock,600);}
      else{setMessage("PINs don't match - try again");setNewPin("");setStep("new1");triggerShake();}
    }
  }
  const keys=[["1","2","3"],["4","5","6"],["7","8","9"],["","0","⌫"]];
  return (
    <div style={{position:"fixed",inset:0,background:"#0a0a08",zIndex:9999,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:'"Courier New",monospace'}}>
      <style>{`
        @keyframes fadeInLock{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
        @keyframes shakeLock{0%{transform:translateX(0)}20%{transform:translateX(-10px)}40%{transform:translateX(10px)}60%{transform:translateX(-8px)}80%{transform:translateX(8px)}100%{transform:translateX(0)}}
        .shake{animation:shakeLock .4s ease}
        .pin-key{background:#141410;border:1px solid #2a2010;color:#e8e0d0;font-size:22px;font-weight:bold;font-family:"Courier New",monospace;width:72px;height:72px;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .12s;user-select:none;-webkit-user-select:none}
        .pin-key:active{background:rgba(200,150,12,.2);border-color:#c8960c;color:#c8960c;transform:scale(.93)}
      `}</style>
      <div style={{marginBottom:32,textAlign:"center"}}>
        <div style={{fontSize:48,marginBottom:10}}>🔐</div>
        <div style={{fontSize:16,fontWeight:"bold",color:"#c8960c",letterSpacing:3}}>RBR KEY MANAGEMENT</div>
        <div style={{fontSize:10,color:"#604020",letterSpacing:2,marginTop:4}}>PROPERTY KEY INVENTORY</div>
      </div>
      <div style={{fontSize:13,color:"#806040",letterSpacing:1,marginBottom:24,minHeight:20}}>{message}</div>
      <div className={shake?"shake":""} style={{display:"flex",gap:18,marginBottom:36}}>
        {[0,1,2,3].map(i=>(
          <div key={i} style={{width:16,height:16,borderRadius:"50%",
            background:i<digits.length?(success?"#50c880":shake?"#e06060":"#c8960c"):"transparent",
            border:`2px solid ${i<digits.length?(success?"#50c880":shake?"#e06060":"#c8960c"):"#3a3020"}`,
            transition:"all .15s"}}/>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,72px)",gap:14}}>
        {keys.flat().map((k,i)=>(
          k===""?<div key={i}/>:
          k==="⌫"?<div key={i} className="pin-key" onClick={()=>setDigits(p=>p.slice(0,-1))} style={{fontSize:20,color:"#806040"}}>⌫</div>:
          <div key={i} className="pin-key" onClick={()=>handleDigit(k)}>{k}</div>
        ))}
      </div>
      {mode==="change"&&onCancel&&(
        <button onClick={onCancel} style={{marginTop:32,background:"transparent",border:"none",color:"#604030",cursor:"pointer",fontSize:12,fontFamily:'"Courier New",monospace',letterSpacing:1}}>Cancel</button>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// AddressInput component
// ─────────────────────────────────────────────────────────────────────────────
interface AddressInputProps {
  value:string; onChange:(v:string,box?:string)=>void;
  onSelect:(addr:string,box:string)=>void; onAddNew:(q:string,box:string)=>void;
  allEntries:AddrEntry[]; boxes:string[]; boxStats:Record<string,BoxStat>;
  activeBox?:string; placeholder?:string;
}
function AddressInput({value,onChange,onSelect,onAddNew,allEntries,boxes,boxStats,activeBox,placeholder}:AddressInputProps){
  const [query,setQuery]=useState(value||"");
  const [open,setOpen]=useState(false);
  const [confirmed,setConfirmed]=useState(!!value);
  const [pickingBox,setPickingBox]=useState(false);
  const ref=useRef<HTMLDivElement>(null);

  useEffect(()=>{
    if(!value){setQuery("");setConfirmed(false);setPickingBox(false);}
    else if(value!==query){setQuery(value);setConfirmed(true);}
  },[value]);
  useEffect(()=>{
    const h=(e:MouseEvent)=>{if(ref.current&&!ref.current.contains(e.target as Node)){setOpen(false);setPickingBox(false);}};
    document.addEventListener("mousedown",h); return()=>document.removeEventListener("mousedown",h);
  },[]);

  const sugs=useMemo(()=>query.length>=2?fuzzySearch(query,allEntries):[],[query,allEntries]);
  const grouped=useMemo(()=>{const g:Record<string,typeof sugs>={};for(const s of sugs){if(!g[s.box])g[s.box]=[];g[s.box].push(s);}return g;},[sugs]);
  const isNew=query.length>=3&&!confirmed&&!allEntries.some(e=>e.addr.toLowerCase()===query.toLowerCase());
  const showDD=open&&(sugs.length>0||isNew);

  function handleChange(e:React.ChangeEvent<HTMLInputElement>){
    const q=e.target.value;setQuery(q);setConfirmed(false);setPickingBox(false);
    onChange("",undefined);setOpen(q.length>=2);
  }
  function pick(addr:string,box:string){setQuery(addr);setConfirmed(true);setOpen(false);setPickingBox(false);onSelect(addr,box);onChange(addr,box);}
  function confirmBox(box:string){if((boxStats[box]?.pct||0)>=100)return;setOpen(false);setPickingBox(false);onAddNew(query,box);}
  function hl(text:string,q:string){
    const i=text.toLowerCase().indexOf(q.toLowerCase());
    if(i<0) return <>{text}</>;
    return <><span>{text.slice(0,i)}</span><mark style={{background:"rgba(200,150,12,.3)",color:"#e8d090",borderRadius:2,padding:"0 1px"}}>{text.slice(i,i+q.length)}</mark><span>{text.slice(i+q.length)}</span></>;
  }
  const icon=confirmed?<span style={{color:"#50c050"}}>✓</span>:isNew?<span style={{color:"#c8960c"}}>✦</span>:query.length>=2?<span style={{color:"#806040"}}>⌕</span>:null;

  return (
    <div ref={ref} style={{position:"relative"}}>
      <div style={{position:"relative"}}>
        <input value={query} onChange={handleChange} onFocus={()=>{if(sugs.length>0)setOpen(true);}}
          placeholder={placeholder||"Type to search all boxes..."} autoComplete="off" autoCorrect="off" spellCheck={false}
          style={{width:"100%",paddingRight:36,borderColor:confirmed?"#3a8a3a":isNew?"#a07020":undefined,transition:"border-color .2s"}}/>
        <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",fontSize:15,pointerEvents:"none"}}>{icon}</div>
      </div>
      {showDD&&(
        <div style={{position:"absolute",top:"calc(100% + 5px)",left:0,right:0,zIndex:9999,background:"#141410",border:"1px solid #3a3020",borderRadius:8,boxShadow:"0 10px 36px rgba(0,0,0,.78)",overflow:"hidden",maxHeight:360,overflowY:"auto"}}>
          {pickingBox?(
            <>
              <div style={{padding:"10px 14px 8px",background:"#0f0f0a",borderBottom:"1px solid #252510",display:"flex",alignItems:"center",gap:8}}>
                <button onClick={()=>setPickingBox(false)} style={{background:"transparent",border:"none",color:"#806040",cursor:"pointer",fontSize:16,padding:0,lineHeight:1}}>←</button>
                <div><div style={{fontSize:12,color:"#c8960c",fontWeight:"bold"}}>Which box?</div><div style={{fontSize:10,color:"#605040",marginTop:1}}>"{query}"</div></div>
              </div>
              <div style={{padding:8}}>
                {boxes.map(box=>{
                  const st=boxStats[box]||{props:0,maxProps:DEFAULT_MAX,pct:0};
                  const full=st.pct>=100;const pc=pctColor(st.pct);
                  return (
                    <div key={box} onClick={()=>!full&&confirmBox(box)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",borderRadius:6,cursor:full?"not-allowed":"pointer",marginBottom:4,border:`1px solid ${full?"#5a2020":"#2a2010"}`,opacity:full?0.65:1}}
                      onMouseEnter={e=>{if(!full)(e.currentTarget as HTMLElement).style.background="rgba(200,150,12,.1)";}}
                      onMouseLeave={e=>{(e.currentTarget as HTMLElement).style.background="transparent";}}>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <span style={{fontSize:20}}>{full?"🔒":"📦"}</span>
                        <div>
                          <div style={{fontSize:13,fontWeight:"bold",color:full?"#e06060":"#e0d0b0"}}>{box}</div>
                          <div style={{fontSize:10,color:full?"#c05050":"#604030",marginTop:2}}>{full?`FULL - ${st.props}/${st.maxProps}`:`${st.props}/${st.maxProps} props · ${st.pct}%`}</div>
                          <div style={{width:110,height:3,background:"#2a2010",borderRadius:2,marginTop:4}}><div style={{width:`${st.pct}%`,height:"100%",background:pc,borderRadius:2}}/></div>
                        </div>
                      </div>
                      <div style={{fontSize:11,color:full?"#c05050":"#c8960c",border:`1px solid ${full?"rgba(200,80,80,.3)":"rgba(200,150,12,.3)"}`,padding:"4px 12px",borderRadius:4,flexShrink:0,marginLeft:12}}>{full?"Full 🔒":"Select →"}</div>
                    </div>
                  );
                })}
              </div>
            </>
          ):(
            <>
              <div style={{padding:"7px 14px 6px",fontSize:9,color:"#806040",letterSpacing:1.5,borderBottom:"1px solid #252510",display:"flex",justifyContent:"space-between",background:"#0f0f0a"}}>
                <span>🔍 ALL BOXES</span><span style={{color:"#504030"}}>{sugs.length} found</span>
              </div>
              {Object.entries(grouped).map(([box,items])=>(
                <div key={box}>
                  <div style={{padding:"5px 14px 3px",fontSize:9,color:"#c8960c",letterSpacing:1,background:"rgba(200,150,12,.05)",borderBottom:"1px solid #1e1e10"}}>
                    📦 {box}{activeBox&&box!==activeBox?<span style={{color:"#806040"}}> (other box)</span>:""}
                  </div>
                  {items.map((s,i)=>(
                    <div key={s.addr} onClick={()=>pick(s.addr,s.box)} style={{padding:"10px 14px",cursor:"pointer",borderBottom:"1px solid #1a1a10",display:"flex",alignItems:"center",justifyContent:"space-between"}}
                      onMouseEnter={e=>(e.currentTarget as HTMLElement).style.background="rgba(200,150,12,.08)"}
                      onMouseLeave={e=>(e.currentTarget as HTMLElement).style.background="transparent"}>
                      <div>
                        <div style={{fontSize:13,color:"#e0d0b0"}}>{hl(s.addr,query)}</div>
                        <div style={{fontSize:10,color:i===0&&Object.keys(grouped)[0]===box?"#c8960c":"#504030",marginTop:2}}>
                          {i===0&&Object.keys(grouped)[0]===box?"⭐ Best match":`Match ${i+1}`}
                          {s.box!==activeBox&&<span style={{color:"#806040",marginLeft:6}}>- in {s.box}</span>}
                        </div>
                      </div>
                      <div style={{fontSize:11,color:"#c8960c",border:"1px solid rgba(200,150,12,.3)",padding:"3px 10px",borderRadius:4,whiteSpace:"nowrap",marginLeft:12,flexShrink:0}}>Use →</div>
                    </div>
                  ))}
                </div>
              ))}
              {isNew&&(
                <div onClick={()=>setPickingBox(true)} style={{padding:"11px 14px",cursor:"pointer",borderTop:sugs.length>0?"2px solid #2a2010":"none",display:"flex",alignItems:"center",justifyContent:"space-between",background:"rgba(200,150,12,.03)"}}
                  onMouseEnter={e=>(e.currentTarget as HTMLElement).style.background="rgba(200,150,12,.09)"}
                  onMouseLeave={e=>(e.currentTarget as HTMLElement).style.background="rgba(200,150,12,.03)"}>
                  <div>
                    <div style={{fontSize:13,color:"#d0a040"}}>➕ Add "{query}" as new property</div>
                    <div style={{fontSize:10,color:"#605040",marginTop:2}}>Not found - tap to choose a box</div>
                  </div>
                  <div style={{fontSize:11,color:"#50c880",border:"1px solid rgba(80,200,128,.35)",padding:"3px 10px",borderRadius:4,whiteSpace:"nowrap",marginLeft:12,flexShrink:0}}>Add New →</div>
                </div>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Main App
// ─────────────────────────────────────────────────────────────────────────────
export default function App(){
  // ── Auth
  const [locked,setLocked]=useState(true);
  const [showChangePin,setShowChangePin]=useState(false);
  const inactivityTimer=useRef<ReturnType<typeof setTimeout>>();

  // ── Firebase sync state
  const [fbReady,setFbReady]=useState(false);   // true once initial load done
  const [syncing,setSyncing]=useState(false);    // true while writing to FB
  const [lastSync,setLastSync]=useState<string>("");
  const [onlineUsers,setOnlineUsers]=useState(1);

  // ── Data
  const [data,setData]=useState<Record<string,PropertyRow[]>>({});
  const [boxSettings,setBoxSettings]=useState<Record<string,BoxSettings>>({});
  const [txLog,setTxLog]=useState<TxLog[]>([]);

  // ── UI state
  const [activeTab,setActiveTab]=useState("inventory");
  const [tabHistory,setTabHistory]=useState<string[]>([]);
  const [activeBox,setActiveBox]=useState("BOX 1");
  const [invSearch,setInvSearch]=useState("");
  const [txAddress,setTxAddress]=useState("");
  const [txBox,setTxBox]=useState("BOX 1");
  const [txKeyType,setTxKeyType]=useState<string>(KEY_TYPES[0]);
  const [txQty,setTxQty]=useState<string>("1");
  const [txMode,setTxMode]=useState<"deposit"|"withdraw">("withdraw");
  const [addAddress,setAddAddress]=useState("");
  const [addBox,setAddBox]=useState("BOX 1");
  const [addLockNo,setAddLockNo]=useState("");
  const [addKeys,setAddKeys]=useState<Record<string,string>>(Object.fromEntries(KEY_TYPES.map(k=>[k,"0"])));
  const [addExisting,setAddExisting]=useState(false);
  const [addAnotherPrompt,setAddAnotherPrompt]=useState(false);
  const [logFilter,setLogFilter]=useState("all");
  const [logSearch,setLogSearch]=useState("");
  const [editRow,setEditRow]=useState<(PropertyRow&{_box:string})|null>(null);
  const [confirmDelete,setConfirmDelete]=useState<{box:string;id:string;address:string}|null>(null);
  const [showNewBox,setShowNewBox]=useState(false);
  const [newBoxName,setNewBoxName]=useState("");
  const [newBoxMax,setNewBoxMax]=useState<string>("100");
  const [editBoxSettings,setEditBoxSettings]=useState<{name:string;maxProps:string}|null>(null);
  const [showClearData,setShowClearData]=useState(false);
  const [clearStep,setClearStep]=useState(1);
  const [clearTarget,setClearTarget]=useState<string>("all");
  const [clearConfirmText,setClearConfirmText]=useState("");
  const [clearPinDigits,setClearPinDigits]=useState<string[]>([]);
  const [clearPinError,setClearPinError]=useState(false);
  const [clearPinShake,setClearPinShake]=useState(false);
  const [toast,setToast]=useState<{msg:string;type:string}|null>(null);
  const toastTimer=useRef<ReturnType<typeof setTimeout>>();
  const importRef=useRef<HTMLInputElement>(null);

  // ── Auto-lock
  useEffect(()=>{
    if(locked) return;
    function resetTimer(){clearTimeout(inactivityTimer.current);inactivityTimer.current=setTimeout(()=>setLocked(true),AUTO_LOCK_MS);}
    const events=["mousedown","touchstart","keydown","scroll"];
    events.forEach(e=>document.addEventListener(e,resetTimer,{passive:true}));
    resetTimer();
    return()=>{clearTimeout(inactivityTimer.current);events.forEach(e=>document.removeEventListener(e,resetTimer));};
  },[locked]);

  // ── Firebase: real-time listeners (start after unlock)
  useEffect(()=>{
    if(locked) return;
    const unsubs:Array<()=>void>=[];

    // Listen to settings
    const settingsUnsub=onSnapshot(collection(db,"settings"),(snap)=>{
      const s:Record<string,BoxSettings>={};
      snap.docs.forEach(d=>{s[d.id]={maxProps:d.data().maxProps||DEFAULT_MAX};});
      setBoxSettings(prev=>{
        // Keep defaults for boxes not yet in Firebase
        const merged={...prev,...s};
        return merged;
      });
    });
    unsubs.push(settingsUnsub);

    // Listen to history (last 200)
    const histUnsub=onSnapshot(
      query(collection(db,"history"),orderBy("createdAt","desc"),limit(200)),
      (snap)=>{
        const logs:TxLog[]=snap.docs.map(d=>{
          const dat=d.data();
          return {id:d.id,ts:dat.ts||"",type:dat.type||"",address:dat.address||"",box:dat.box||"",keyType:dat.keyType||"",qty:dat.qty||0,user:dat.user||""};
        });
        setTxLog(logs);
      }
    );
    unsubs.push(histUnsub);

    // Load boxes from Firebase; if empty, seed with initial data
    (async()=>{
      const boxSnap=await getDocs(collection(db,"boxes"));
      if(boxSnap.empty){
        // First time - upload seed data
        setSyncing(true);
        showToastFn("Uploading initial data to Firebase...","success");
        const batch=writeBatch(db);
        for(const [box,rows] of Object.entries(SEED_DATA)){
          for(const r of rows as any[]){
            const id=r._id||`${box}-${Math.random()}`;
            const ref=doc(db,"boxes",safeId(box),"properties",safeId(id));
            batch.set(ref,{lockNo:r["Lock Box No."]||0,address:r["Property Address"]||"",mainDoor:r["Main Door Key"]||0,mailBox:r["Mail Box Key"]||0,penDrive:r["Pen drive"]||0,smartKey:r["Smart Key"]||0,otherKey:r["Other Key"]||0,_id:id,updatedAt:serverTimestamp()});
          }
          // default settings
          const sref=doc(db,"settings",safeId(box));
          batch.set(sref,{maxProps:DEFAULT_MAX,updatedAt:serverTimestamp()},{mergeFields:["maxProps"]});
        }
        await batch.commit();
        setSyncing(false);
        showToastFn("✓ Data uploaded to Firebase","success");
      }

      // Now set up per-box listeners
      const loadedBoxes=boxSnap.empty?Object.keys(SEED_DATA):boxSnap.docs.map(d=>d.id);
      // Also listen for dynamically added boxes
      const knownBoxes=new Set(loadedBoxes);

      for(const box of loadedBoxes){
        const unsub=onSnapshot(collection(db,"boxes",safeId(box),"properties"),(snap)=>{
          const rows:PropertyRow[]=snap.docs.map(d=>fbRowToProperty(d.data(), d.id));
          rows.sort((a,b)=>a["Lock Box No."]-b["Lock Box No."]);
          setData(prev=>({...prev,[box]:rows}));
          setLastSync(nowStr());
          // Clear tx/add address if the property was deleted from this box
          setTxAddress(prev=>{
            if(prev && box===txBox && !rows.some(r=>r["Property Address"]===prev)){
              return ""; // property deleted - clear selection
            }
            return prev;
          });
        });
        unsubs.push(unsub);
        if(!boxSettings[box]) setBoxSettings(prev=>({...prev,[box]:{maxProps:DEFAULT_MAX}}));
      }
      setActiveBox(loadedBoxes[0]||"BOX 1");
      setTxBox(loadedBoxes[0]||"BOX 1");
      setAddBox(loadedBoxes[0]||"BOX 1");
      setFbReady(true);
    })();

    return()=>unsubs.forEach(u=>u());
  },[locked]);

  // ── Toast helper (defined before use in effect)
  function showToastFn(msg:string,type="success"){
    setToast({msg,type});clearTimeout(toastTimer.current);
    toastTimer.current=setTimeout(()=>setToast(null),3800);
  }
  const showToast=useCallback(showToastFn,[]);

  function navigateTo(tab:string){setTabHistory(prev=>[...prev,activeTab]);setActiveTab(tab);}
  function goBack(){setTabHistory(prev=>{const h=[...prev];const last=h.pop()||"inventory";setActiveTab(last);return h;});}

  function addLog(type:string,address:string,box:string,keyType:string,qty:number){
    const log:TxLog={id:`log-${Date.now()}`,ts:nowStr(),type,address,box,keyType,qty};
    fbSaveLog(log).catch(()=>{});
    // optimistic local update
    setTxLog(prev=>[log,...prev]);
  }

  // ── Derived
  const boxes=useMemo(()=>Object.keys(data),[data]);
  const boxStats=useMemo(()=>getBoxStats(data,boxSettings),[data,boxSettings]);
  const entries=useMemo(()=>Object.entries(data).flatMap(([box,rows])=>rows.map(r=>({addr:r["Property Address"],box}))),[data]);

  // When data changes, validate current tx selection - clear if deleted
  useEffect(()=>{
    if(!txAddress) return;
    const existsInAnyBox=entries.some(e=>e.addr===txAddress);
    if(!existsInAnyBox){
      setTxAddress("");
      showToastFn(`Property was deleted - selection cleared`,"error");
    } else {
      // If not in selected box, auto-switch to correct box
      const correctEntry=entries.find(e=>e.addr===txAddress);
      if(correctEntry && correctEntry.box!==txBox){
        setTxBox(correctEntry.box);
      }
    }
  },[data, txAddress]);
  const filteredRows=useMemo(()=>{
    const rows=data[activeBox]||[];
    if(!invSearch.trim()) return rows;
    const q=invSearch.toLowerCase();
    return rows.filter(r=>r["Property Address"].toLowerCase().includes(q)||String(r["Lock Box No."]).includes(q));
  },[data,activeBox,invSearch]);
  const filteredLog=useMemo(()=>{
    let log=txLog;
    if(logFilter!=="all") log=log.filter(l=>l.type===logFilter);
    if(logSearch.trim()){const q=logSearch.toLowerCase();log=log.filter(l=>l.address.toLowerCase().includes(q)||l.box.toLowerCase().includes(q)||l.keyType.toLowerCase().includes(q));}
    return log;
  },[txLog,logFilter,logSearch]);
  const activeBoxStat=boxStats[activeBox]||{props:0,keys:0,maxProps:DEFAULT_MAX,pct:0};
  const currentTxRow=txAddress?(data[txBox]||[]).find(r=>r["Property Address"]===txAddress):null;

  // ── Handlers
  async function handleTransaction(){
    const qty=parseInt(String(txQty))||0;
    if(!txAddress){showToast("Select a property address","error");return;}
    if(qty<=0){showToast("Enter a valid quantity","error");return;}
    const idx=(data[txBox]||[]).findIndex(r=>r["Property Address"]===txAddress);
    if(idx===-1){
      showToast(`"${txAddress}" not found in ${txBox} - it may have been deleted`,"error");
      setTxAddress(""); // clear the invalid selection
      return;
    }
    const row=data[txBox][idx];
    const current=row[txKeyType as KeyType]||0;
    if(txMode==="withdraw"&&current<qty){showToast(`Only ${current} ${txKeyType}(s) available`,"error");return;}
    if(txMode==="deposit"&&(boxStats[txBox]?.pct||0)>=100){showToast(`⚠ ${txBox} is FULL`,"error");return;}
    const updated={...row,[txKeyType]:txMode==="withdraw"?current-qty:current+qty};
    // optimistic local update
    const updatedRows=[...data[txBox]];updatedRows[idx]=updated;
    setData(prev=>({...prev,[txBox]:updatedRows}));
    // sync to Firebase
    setSyncing(true);
    try{await fbSaveProperty(txBox,updated);}catch{showToast("Sync error - check connection","error");}
    setSyncing(false);
    addLog(txMode,txAddress,txBox,txKeyType,qty);
    showToast(`${txMode==="withdraw"?"⬆ Withdrawn":"⬇ Deposited"} ${qty}× ${txKeyType} - ${txBox}`);
    setTxQty("1");
  }

  async function handleAddNew(){
    const addr=addAddress.trim();
    if(!addr){showToast("Address is required","error");return;}
    const qty=KEY_TYPES.reduce((s,k)=>s+(parseInt(String(addKeys[k]))||0),0);
    const existIdx=(data[addBox]||[]).findIndex(r=>r["Property Address"].toLowerCase()===addr.toLowerCase());
    setSyncing(true);
    if(existIdx!==-1){
      if(qty===0){showToast("Enter at least 1 key to add","error");setSyncing(false);return;}
      const row=data[addBox][existIdx];
      const updated={...row};
      KEY_TYPES.forEach(k=>{updated[k]=(row[k]||0)+(parseInt(String(addKeys[k]))||0);});
      const rows=[...data[addBox]];rows[existIdx]=updated;
      setData(prev=>({...prev,[addBox]:rows}));
      try{await fbSaveProperty(addBox,updated);}catch{showToast("Sync error","error");}
      addLog("deposit",addr,addBox,"(multiple)",qty);
      showToast(`✓ Added ${qty} key(s) to "${addr}"`);
    } else {
      if((boxStats[addBox]?.pct||0)>=100){showToast(`⚠ ${addBox} is FULL`,"error");setSyncing(false);return;}
      const nextNo=Math.max(...(data[addBox]||[]).map(r=>r["Lock Box No."]),0)+1;
      const newRow:PropertyRow={"Lock Box No.":addLockNo?parseInt(addLockNo):nextNo,"Property Address":addr,_id:`${addBox}-new-${Date.now()}`,...Object.fromEntries(KEY_TYPES.map(k=>[k,parseInt(String(addKeys[k]))||0])) as any};
      setData(prev=>({...prev,[addBox]:[...(prev[addBox]||[]),newRow]}));
      try{await fbSaveProperty(addBox,newRow);}catch{showToast("Sync error","error");}
      addLog("add",addr,addBox,"(new property)",qty);
      showToast(`✓ Added "${addr}" to ${addBox}`);
    }
    setSyncing(false);
    setAddAddress("");setAddLockNo("");setAddExisting(false);
    setAddKeys(Object.fromEntries(KEY_TYPES.map(k=>[k,"0"])));
    setAddAnotherPrompt(true);setTimeout(()=>setAddAnotherPrompt(false),6000);
  }

  function handleDelete(box:string,id:string){
    const row=(data[box]||[]).find(r=>r._id===id);
    setConfirmDelete({box,id,address:row?.["Property Address"]||"this property"});
  }
  async function confirmDeleteNow(){
    if(!confirmDelete) return;
    const {box,id,address}=confirmDelete;
    // optimistic UI update first
    setData(prev=>({...prev,[box]:prev[box].filter(r=>r._id!==id)}));
    if(editRow?._id===id) setEditRow(null);
    setConfirmDelete(null);
    setSyncing(true);
    let deleted=false;
    // Try 1: delete by exact document id
    try{
      await deleteDoc(doc(db,"boxes",safeId(box),"properties",id));
      deleted=true;
    }catch(e1:any){
      showToast(`Error: ${e1?.code||e1?.message||"unknown"}`, "error");
    }
    // Try 2: query by address and delete all matching docs
    if(!deleted){
      try{
        const q=query(collection(db,"boxes",safeId(box),"properties"),where("address","==",address));
        const snap=await getDocs(q);
        for(const d of snap.docs){ await deleteDoc(d.ref); deleted=true; }
      }catch(e2:any){
        showToast(`Error2: ${e2?.code||e2?.message||"unknown"}`, "error");
      }
    }
    if(deleted) showToast("Property deleted");
    else showToast("Could not delete - check Firebase Rules", "error");
    setSyncing(false);
  }

  async function handleSaveEdit(){
    if(!editRow) return;
    setData(prev=>({...prev,[editRow._box]:prev[editRow._box].map(r=>r._id===editRow._id?{...editRow}:r)}));
    setSyncing(true);
    try{await fbSaveProperty(editRow._box,editRow);}catch{showToast("Sync error","error");}
    setSyncing(false);
    setEditRow(null);showToast("Record updated");
  }

  async function handleAddBox(){
    const name=newBoxName.trim().toUpperCase();
    if(!name){showToast("Box name required","error");return;}
    if(data[name]){showToast(`"${name}" already exists`,"error");return;}
    const maxP=parseInt(String(newBoxMax))||DEFAULT_MAX;
    setData(prev=>({...prev,[name]:[]}));
    setBoxSettings(prev=>({...prev,[name]:{maxProps:maxP}}));
    setSyncing(true);
    try{await fbSaveBoxSettings(name,maxP);}catch{}
    setSyncing(false);
    setActiveBox(name);setShowNewBox(false);setNewBoxName("");setNewBoxMax("100");
    showToast(`✓ Created ${name}`);
  }

  async function handleSaveBoxSettings(){
    if(!editBoxSettings) return;
    const max=parseInt(String(editBoxSettings.maxProps))||DEFAULT_MAX;
    const curr=boxStats[editBoxSettings.name]?.props||0;
    if(max<curr){showToast(`⚠ Max (${max}) below current count (${curr})`,"error");return;}
    setBoxSettings(prev=>({...prev,[editBoxSettings.name]:{maxProps:max}}));
    setSyncing(true);
    try{await fbSaveBoxSettings(editBoxSettings.name,max);}catch{}
    setSyncing(false);
    setEditBoxSettings(null);showToast(`✓ Updated ${editBoxSettings.name}`);
  }

  async function handleClearData(){
    setSyncing(true);
    if(clearTarget==="all"){
      const empty:Record<string,PropertyRow[]>={};
      for(const box of Object.keys(data)) empty[box]=[];
      setData(empty);setTxLog([]);
      try{await Promise.all(Object.keys(data).map(b=>fbClearBox(b)));await fbClearHistory();}catch{}
      showToast("✓ All data cleared");
    } else if(clearTarget==="history"){
      setTxLog([]);try{await fbClearHistory();}catch{};showToast("✓ History cleared");
    } else {
      setData(prev=>({...prev,[clearTarget]:[]}));
      try{await fbClearBox(clearTarget);}catch{};showToast(`✓ ${clearTarget} cleared`);
    }
    setSyncing(false);
    setShowClearData(false);setClearConfirmText("");setClearStep(1);setClearPinDigits([]);setClearPinError(false);
    navigateTo("inventory");
  }

  function handleClearPinDigit(d:string){
    if(clearPinDigits.length>=4) return;
    const next=[...clearPinDigits,d];setClearPinDigits(next);
    if(next.length===4){
      setTimeout(()=>{
        if(next.join("")===getStoredPasscode()){setClearPinDigits([]);setClearPinError(false);setClearStep(3);}
        else{setClearPinShake(true);setTimeout(()=>{setClearPinShake(false);setClearPinDigits([]);setClearPinError(true);},500);}
      },120);
    }
  }

  function handleImport(e:React.ChangeEvent<HTMLInputElement>){
    const file=e.target.files?.[0];if(!file) return;
    const reader=new FileReader();
    reader.onload=async(evt)=>{
      try{
        const wb=XLSX.read(evt.target?.result,{type:"binary"});
        let totalImported=0;
        setSyncing(true);
        for(const sheetName of wb.SheetNames){
          if(sheetName.toLowerCase()==="history") continue;
          const ws=wb.Sheets[sheetName];
          const rows:any[]=XLSX.utils.sheet_to_json(ws,{defval:0});
          if(rows.length===0) continue;
          const imported:PropertyRow[]=rows.map((r:any,i:number)=>({
            "Lock Box No.":parseInt(r["Lock Box No."])||i+1,
            "Property Address":String(r["Property Address"]||"").trim(),
            "Main Door Key":parseInt(r["Main Door Key"])||0,"Mail Box Key":parseInt(r["Mail Box Key"])||0,
            "Pen drive":parseInt(r["Pen drive"])||0,"Smart Key":parseInt(r["Smart Key"])||0,"Other Key":parseInt(r["Other Key"])||0,
            _id:`${sheetName}-import-${i}-${Date.now()}`
          })).filter(r=>r["Property Address"].length>0);
          setData(prev=>({...prev,[sheetName]:imported}));
          setBoxSettings(prev=>prev[sheetName]?prev:{...prev,[sheetName]:{maxProps:DEFAULT_MAX}});
          // sync to Firebase
          await fbClearBox(sheetName); // fbClearBox already uses safeId internally
          const batch=writeBatch(db);
          imported.forEach(r=>{const ref=doc(db,"boxes",sheetName,"properties",safeId(r._id));batch.set(ref,{lockNo:r["Lock Box No."],address:r["Property Address"],mainDoor:r["Main Door Key"],mailBox:r["Mail Box Key"],penDrive:r["Pen drive"],smartKey:r["Smart Key"],otherKey:r["Other Key"],_id:r._id,updatedAt:serverTimestamp()});});
          await batch.commit();
          totalImported+=imported.length;
        }
        setSyncing(false);
        showToast(`✓ Imported ${totalImported} properties - synced to Firebase`);
        setActiveTab("inventory");
      }catch{showToast("Failed to read Excel file","error");setSyncing(false);}
    };
    reader.readAsBinaryString(file);e.target.value="";
  }

  function exportXLSX(){
    const wb=XLSX.utils.book_new();
    for(const [boxName,rows] of Object.entries(data)){
      const ws=XLSX.utils.json_to_sheet(rows.map(r=>({"Lock Box No.":r["Lock Box No."],"Property Address":r["Property Address"],"Main Door Key":r["Main Door Key"],"Mail Box Key":r["Mail Box Key"],"Pen drive":r["Pen drive"],"Smart Key":r["Smart Key"],"Other Key":r["Other Key"]})));
      ws["!cols"]=[{wch:14},{wch:44},{wch:14},{wch:14},{wch:12},{wch:12},{wch:12}];
      XLSX.utils.book_append_sheet(wb,ws,boxName);
    }
    if(txLog.length>0){
      const ws2=XLSX.utils.json_to_sheet(txLog.map(l=>({Date:l.ts,Type:l.type.toUpperCase(),Address:l.address,"Lock Box":l.box,"Key Type":l.keyType,Qty:l.qty})));
      XLSX.utils.book_append_sheet(wb,ws2,"History");
    }
    XLSX.writeFile(wb,"RBR-Key_List_Updated.xlsx");showToast("Excel downloaded!");
  }

  // ── CSS
  const CSS=`
    *{box-sizing:border-box;margin:0;padding:0}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#1a1a1a}::-webkit-scrollbar-thumb{background:#c8960c;border-radius:3px}
    input,select{background:#1c1c1c;border:1px solid #3a3020;color:#e8e0d0;font-family:"Courier New",monospace;border-radius:4px;padding:8px 12px;font-size:13px;outline:none;transition:border-color .2s}
    input:focus,select:focus{border-color:#c8960c;box-shadow:0 0 0 2px rgba(200,150,12,.15)}
    select option{background:#1c1c1c}
    table{border-collapse:collapse;width:100%}
    th{background:#1a1500;color:#c8960c;font-size:11px;letter-spacing:1px;text-transform:uppercase;padding:10px 12px;border-bottom:2px solid #3a3020;white-space:nowrap}
    td{padding:9px 12px;border-bottom:1px solid #1e1e1e;font-size:13px}
    tr:hover td{background:rgba(200,150,12,.04)}
    .btn{cursor:pointer;border:none;border-radius:4px;font-family:"Courier New",monospace;font-weight:bold;letter-spacing:.5px;transition:all .15s}
    .btn-amber{background:#c8960c;color:#0f0f0f;padding:9px 20px}.btn-amber:hover{background:#e8aa0e}
    .btn-outline{background:transparent;border:1px solid #3a3020;color:#a09070;padding:8px 16px;font-size:12px}.btn-outline:hover{border-color:#c8960c;color:#c8960c}
    .btn-ghost{background:transparent;border:1px solid #2a2010;color:#706050;padding:5px 10px;font-size:11px}.btn-ghost:hover{border-color:#c8960c;color:#c8960c}
    .btn-danger{background:transparent;border:1px solid #5a2020;color:#c05050;padding:5px 10px;font-size:11px}.btn-danger:hover{background:#5a2020;color:#ffaaaa}
    .btn-edit{background:transparent;border:1px solid #203050;color:#5090c0;padding:5px 10px;font-size:11px}.btn-edit:hover{background:#203050;color:#90c0ff}
    .badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:bold}
    .badge-a{background:rgba(200,150,12,.15);color:#c8960c;border:1px solid rgba(200,150,12,.3)}
    .badge-z{background:rgba(80,80,80,.2);color:#505050;border:1px solid #252525}
    .badge-h{background:rgba(80,200,120,.1);color:#50c880;border:1px solid rgba(80,200,120,.25)}
    .badge-r{background:rgba(200,80,80,.1);color:#e07070;border:1px solid rgba(200,80,80,.25)}
    .badge-dep{background:rgba(80,160,80,.15);color:#70d070;border:1px solid rgba(80,160,80,.3)}
    .badge-wit{background:rgba(200,100,80,.15);color:#e09070;border:1px solid rgba(200,100,80,.3)}
    .badge-add{background:rgba(100,140,200,.15);color:#90b8e8;border:1px solid rgba(100,140,200,.3)}
    .sc{background:#141410;border:1px solid #2a2510;border-radius:8px;padding:15px 18px}
    .fg{display:grid;gap:13px}
    .lbl{font-size:11px;color:#806040;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px}
    .sec{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#806040;margin-bottom:16px;border-left:3px solid #c8960c;padding-left:10px}
    .mo{position:fixed;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:center;justify-content:center;z-index:1000}
    .md{background:#141410;border:1px solid #3a3020;border-radius:10px;padding:26px;width:520px;max-width:95vw;max-height:90vh;overflow-y:auto}
    .ni{padding:10px 18px;cursor:pointer;display:flex;align-items:center;gap:10px;font-size:13px;transition:all .15s}
    .pb{height:4px;border-radius:2px;background:#2a2010;overflow:hidden;margin-top:6px}
    .pf{height:100%;border-radius:2px;transition:width .3s}
    .shake{animation:shakeLock .4s ease}
    @keyframes shakeLock{0%{transform:translateX(0)}20%{transform:translateX(-10px)}40%{transform:translateX(10px)}60%{transform:translateX(-8px)}80%{transform:translateX(8px)}100%{transform:translateX(0)}}
    @keyframes ddFade{from{opacity:0;transform:translateY(-5px)}to{opacity:1;transform:translateY(0)}}
    @keyframes slideIn{from{opacity:0;transform:translateX(20px)}to{opacity:1;transform:translateX(0)}}
    @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
    .syncing{animation:pulse 1s ease infinite}
  `;

  // ── Loading screen
  if(locked) return <LockScreen onUnlock={()=>setLocked(false)}/>;
  if(showChangePin) return <LockScreen mode="change" onUnlock={()=>setShowChangePin(false)} onCancel={()=>setShowChangePin(false)}/>;

  if(!fbReady) return (
    <div style={{background:"#0f0f0f",minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:'"Courier New",monospace',color:"#e8e0d0"}}>
      <div style={{fontSize:40,marginBottom:20}}>🔐</div>
      <div style={{fontSize:14,color:"#c8960c",letterSpacing:2,marginBottom:8}}>RBR KEY MANAGEMENT</div>
      <div style={{fontSize:12,color:"#806040",marginBottom:24}}>Connecting to Firebase...</div>
      <div style={{display:"flex",gap:8}}>
        {[0,1,2].map(i=>(<div key={i} style={{width:8,height:8,borderRadius:"50%",background:"#c8960c",animation:`pulse 1s ease ${i*0.2}s infinite`}}/>))}
      </div>
    </div>
  );

  return (
    <div style={{fontFamily:'"Courier New",monospace',background:"#0f0f0f",minHeight:"100vh",color:"#e8e0d0"}}>
      <style>{CSS}</style>

      {/* Header */}
      <div style={{background:"#0a0a08",borderBottom:"2px solid #2a2010",padding:"10px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
          {activeTab!=="inventory"&&(
            <button onClick={goBack} style={{background:"transparent",border:"1px solid #3a3020",color:"#c8960c",borderRadius:6,padding:"6px 10px",cursor:"pointer",fontSize:13,fontFamily:'"Courier New",monospace',fontWeight:"bold"}}>← Back</button>
          )}
          <button onClick={()=>{setTabHistory([]);setActiveTab("inventory");}} style={{background:activeTab==="inventory"?"rgba(200,150,12,.15)":"transparent",border:"1px solid #3a3020",color:activeTab==="inventory"?"#c8960c":"#806040",borderRadius:6,padding:"6px 10px",cursor:"pointer",fontSize:14,fontFamily:'"Courier New",monospace'}}>🏠</button>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,flex:1,justifyContent:"center"}}>
          {syncing&&<span className="syncing" style={{fontSize:10,color:"#c8960c"}}>⟳ Syncing</span>}
          {!syncing&&lastSync&&<span style={{fontSize:9,color:"#404030"}}>✓ {lastSync}</span>}
        </div>
        <div style={{display:"flex",gap:6,flexShrink:0}}>
          <input ref={importRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleImport}/>
          <button className="btn btn-outline" onClick={()=>importRef.current?.click()} style={{fontSize:10,padding:"6px 8px",color:"#50c880",borderColor:"#207020"}}>⬆ Import</button>
          <button className="btn btn-outline" onClick={()=>setShowNewBox(true)} style={{fontSize:10,padding:"6px 8px"}}>➕</button>
          <button className="btn btn-outline" onClick={()=>setShowChangePin(true)} style={{fontSize:10,padding:"6px 8px"}}>🔒</button>
          <button className="btn btn-amber" onClick={exportXLSX} style={{fontSize:10,padding:"6px 8px"}}>⬇</button>
          <button onClick={()=>{setShowClearData(true);setClearStep(1);setClearConfirmText("");setClearPinDigits([]);setClearPinError(false);}} style={{background:"transparent",border:"1px solid #5a2020",color:"#c05050",borderRadius:4,padding:"6px 8px",fontSize:10,cursor:"pointer",fontFamily:'"Courier New",monospace',fontWeight:"bold"}}>🗑</button>
        </div>
      </div>

      {/* Stats */}
      <div style={{background:"#0c0c0a",borderBottom:"1px solid #1e1e10",padding:"10px 14px",display:"flex",gap:10,overflowX:"auto"}}>
        <div className="sc" style={{flexShrink:0,minWidth:90}}>
          <div style={{fontSize:18,fontWeight:"bold",color:"#c8960c"}}>{Object.values(data).flat().length}</div>
          <div style={{fontSize:9,color:"#806040",letterSpacing:1,marginTop:2}}>PROPERTIES</div>
        </div>
        <div className="sc" style={{flexShrink:0,minWidth:90}}>
          <div style={{fontSize:18,fontWeight:"bold",color:"#90c870"}}>{Object.values(data).flat().reduce((s,r)=>s+totalKeys(r),0)}</div>
          <div style={{fontSize:9,color:"#806040",letterSpacing:1,marginTop:2}}>TOTAL KEYS</div>
        </div>
        {Object.entries(boxStats).map(([box,s])=>{
          const pc=pctColor(s.pct);
          return (
            <div key={box} className="sc" style={{flexShrink:0,minWidth:110,cursor:"pointer"}} onClick={()=>{setActiveBox(box);navigateTo("inventory");}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{fontSize:18,fontWeight:"bold",color:s.pct>=100?"#e06060":"#7090d0"}}>{s.props}</div>
                <div style={{fontSize:9,color:pc,background:"rgba(200,150,12,.08)",padding:"2px 5px",borderRadius:4}}>{s.pct}%</div>
              </div>
              <div style={{fontSize:9,color:"#806040",letterSpacing:1,marginTop:2}}>{box}{s.pct>=100?" 🔒":""}</div>
              <div style={{fontSize:9,color:"#504030"}}>{s.props}/{s.maxProps} props</div>
              <div className="pb"><div className="pf" style={{width:`${s.pct}%`,background:pc}}/></div>
            </div>
          );
        })}
        <div className="sc" style={{flexShrink:0,minWidth:90}}>
          <div style={{fontSize:18,fontWeight:"bold",color:"#7080c0"}}>{txLog.length}</div>
          <div style={{fontSize:9,color:"#806040",letterSpacing:1,marginTop:2}}>TRANSACTIONS</div>
        </div>
      </div>

      <div style={{display:"flex",minHeight:"calc(100vh - 140px)"}}>
        {/* Sidebar */}
        <div style={{width:200,background:"#0c0c0a",borderRight:"1px solid #1e1e10",padding:"14px 0",flexShrink:0}}>
          <div style={{padding:"0 14px 8px",fontSize:10,color:"#504030",letterSpacing:2}}>NAVIGATE</div>
          {[["inventory","📋","Inventory"],["enter","⬇","Deposit"],["withdraw","⬆","Withdraw"],["add","➕","Add Property"],["history","📜","History"]].map(([id,icon,label])=>(
            <div key={id} className="ni" onClick={()=>{navigateTo(id);if(id==="enter")setTxMode("deposit");if(id==="withdraw")setTxMode("withdraw");}}
              style={{background:activeTab===id?"rgba(200,150,12,.08)":"transparent",borderLeft:activeTab===id?"3px solid #c8960c":"3px solid transparent",color:activeTab===id?"#c8960c":"#706050"}}>
              <span>{icon}</span>{label}
              {id==="history"&&txLog.length>0&&<span style={{marginLeft:"auto",fontSize:10,background:"rgba(200,150,12,.2)",color:"#c8960c",padding:"1px 5px",borderRadius:8}}>{txLog.length}</span>}
            </div>
          ))}
          <div style={{padding:"14px 14px 6px",fontSize:10,color:"#504030",letterSpacing:2,display:"flex",justifyContent:"space-between",alignItems:"center",paddingRight:8}}>
            <span>BOXES</span>
            <button className="btn" onClick={()=>setShowNewBox(true)} style={{background:"transparent",border:"1px solid #3a3020",color:"#806040",padding:"2px 6px",fontSize:10,borderRadius:4,cursor:"pointer"}}>+</button>
          </div>
          {boxes.map(box=>{
            const bs=boxStats[box]||{props:0,keys:0,maxProps:DEFAULT_MAX,pct:0};
            const pc=pctColor(bs.pct);
            return (
              <div key={box} style={{borderLeft:activeBox===box?"3px solid #c8960c":"3px solid transparent"}}>
                <div className="ni" onClick={()=>setActiveBox(box)} style={{background:activeBox===box?"rgba(200,150,12,.05)":"transparent",color:activeBox===box?"#e0c080":"#706050",justifyContent:"space-between",paddingRight:8}}>
                  <span>{bs.pct>=100?"🔒":"📦"} {box}</span>
                  <div style={{display:"flex",alignItems:"center",gap:5}}>
                    <span style={{fontSize:10,color:pc}}>{bs.pct}%</span>
                    <button className="btn" onClick={e=>{e.stopPropagation();setEditBoxSettings({name:box,maxProps:String(bs.maxProps||DEFAULT_MAX)});}} style={{background:"transparent",border:"none",color:"#504030",fontSize:11,cursor:"pointer",padding:"0 2px"}}>⚙</button>
                  </div>
                </div>
                <div style={{padding:"0 16px 6px"}}>
                  <div className="pb" style={{marginTop:0}}><div className="pf" style={{width:`${bs.pct}%`,background:pc}}/></div>
                  <div style={{fontSize:9,color:"#504030",marginTop:2}}>{bs.props}/{bs.maxProps} props · {bs.keys} keys</div>
                </div>
              </div>
            );
          })}
          <div style={{padding:"12px 14px 6px",fontSize:10,color:"#504030",letterSpacing:2}}>KEY TYPES</div>
          {KEY_TYPES.map(k=>{
            const total=Object.values(data).flat().reduce((s,r)=>s+(r[k]||0),0);
            return (<div key={k} style={{padding:"3px 16px",display:"flex",justifyContent:"space-between",fontSize:11,color:"#706050"}}><span>{KEY_ICONS[k]} {k.replace(" Key","")}</span><span style={{color:"#c8960c"}}>{total}</span></div>);
          })}
        </div>

        {/* Main */}
        <div style={{flex:1,padding:"20px 22px",overflow:"auto"}}>

          {/* INVENTORY */}
          {activeTab==="inventory"&&(
            <div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div className="sec" style={{margin:0}}>{activeBox} - {filteredRows.length} Properties</div>
                  {activeBoxStat.pct>=80&&<span className={`badge ${activeBoxStat.pct>=100?"badge-r":"badge-a"}`}>{activeBoxStat.pct}%{activeBoxStat.pct>=100?" 🔒":""}</span>}
                </div>
                <input placeholder="Search..." value={invSearch} onChange={e=>setInvSearch(e.target.value)} style={{width:220}}/>
              </div>
              <div style={{background:"#0a0a08",border:"1px solid #1e1e10",borderRadius:6,padding:"9px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:12}}>
                <div style={{fontSize:11,color:"#806040",whiteSpace:"nowrap"}}>Capacity</div>
                <div style={{flex:1}}><div className="pb" style={{height:5}}><div className="pf" style={{width:`${activeBoxStat.pct}%`,background:pctColor(activeBoxStat.pct)}}/></div></div>
                <div style={{fontSize:11,color:pctColor(activeBoxStat.pct),whiteSpace:"nowrap"}}>{activeBoxStat.props}/{activeBoxStat.maxProps}{activeBoxStat.pct>=100?" FULL 🔒":""}</div>
                <button className="btn btn-ghost" onClick={()=>setEditBoxSettings({name:activeBox,maxProps:String(activeBoxStat.maxProps||DEFAULT_MAX)})}>⚙</button>
              </div>
              <div style={{background:"#0c0c0a",border:"1px solid #1e1e10",borderRadius:8,overflow:"hidden"}}>
                <div style={{overflowX:"auto"}}>
                  <table>
                    <thead><tr><th>#</th><th>Property Address</th>{KEY_TYPES.map(k=><th key={k}>{KEY_ICONS[k]} {k}</th>)}<th>Total</th><th></th></tr></thead>
                    <tbody>
                      {filteredRows.map(row=>{
                        const tot=totalKeys(row);
                        return (
                          <tr key={row._id}>
                            <td style={{color:"#605040",fontSize:11}}>{row["Lock Box No."]}</td>
                            <td style={{color:"#e0d0b0",maxWidth:240,fontSize:12}}>{row["Property Address"]}</td>
                            {KEY_TYPES.map(k=>(<td key={k} style={{textAlign:"center"}}>{row[k]>0?<span className={`badge ${row[k]>=3?"badge-h":"badge-a"}`}>{row[k]}</span>:<span className="badge badge-z">-</span>}</td>))}
                            <td style={{textAlign:"center"}}><span className={`badge ${tot>0?"badge-a":"badge-z"}`}>{tot}</span></td>
                            <td><div style={{display:"flex",gap:5}}>
                              <button className="btn btn-edit" onClick={()=>setEditRow({...row,_box:activeBox})}>Edit</button>
                              <button className="btn btn-danger" onClick={()=>handleDelete(activeBox,row._id)}>🗑</button>
                            </div></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* DEPOSIT / WITHDRAW */}
          {(activeTab==="enter"||activeTab==="withdraw")&&(
            <div style={{maxWidth:520}}>
              <div className="sec">Key Transaction</div>
              <div style={{display:"flex",border:"1px solid #3a3020",borderRadius:6,overflow:"hidden",marginBottom:20}}>
                <button className="btn" onClick={()=>{setTxMode("deposit");navigateTo("enter");}} style={{flex:1,padding:"9px",fontSize:12,letterSpacing:1,textTransform:"uppercase",background:txMode==="deposit"?"#50a050":"#081a08",color:txMode==="deposit"?"#fff":"#50c050"}}>⬇ Deposit</button>
                <button className="btn" onClick={()=>{setTxMode("withdraw");navigateTo("withdraw");}} style={{flex:1,padding:"9px",fontSize:12,letterSpacing:1,textTransform:"uppercase",background:txMode==="withdraw"?"#c05050":"#1a0808",color:txMode==="withdraw"?"#fff":"#c05050"}}>⬆ Withdraw</button>
              </div>
              <div className="fg">
                <div><div className="lbl">Property Address</div>
                  <AddressInput value={txAddress} onChange={(v,box)=>{setTxAddress(v);if(box)setTxBox(box);}} onSelect={(a,b)=>{setTxAddress(a);setTxBox(b);}}
                    onAddNew={(q,box)=>{setAddAddress(q);setAddBox(box);setAddExisting(false);navigateTo("add");showToast(`Switched to Add - pre-filled in ${box}`);}}
                    allEntries={entries} boxes={boxes} boxStats={boxStats} activeBox={activeBox} placeholder="Type address - all boxes..."/>
                </div>
                {txAddress&&(
                  <div style={{background:"#0a0a08",border:"1px solid #2a2010",borderRadius:6,padding:"12px 14px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                      <div className="lbl" style={{margin:0}}>Box</div>
                      <select value={txBox} onChange={e=>setTxBox(e.target.value)} style={{padding:"4px 8px",fontSize:12}}>{boxes.map(b=><option key={b}>{b}</option>)}</select>
                    </div>
                    {currentTxRow?(
                      <>
                        <div className="lbl" style={{marginBottom:6}}>Lock Box #{currentTxRow["Lock Box No."]}</div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                          {KEY_TYPES.map(k=>(<div key={k} style={{fontSize:12,color:(currentTxRow[k]||0)>0?"#c8960c":"#404040"}}>{KEY_ICONS[k]} {k}: <strong>{currentTxRow[k]||0}</strong></div>))}
                        </div>
                      </>
                    ):<div style={{fontSize:12,color:"#604030"}}>Not found in {txBox}</div>}
                  </div>
                )}
                <div><div className="lbl">Key Type</div><select value={txKeyType} onChange={e=>setTxKeyType(e.target.value)} style={{width:"100%"}}>{KEY_TYPES.map(k=><option key={k}>{k}</option>)}</select></div>
                <div><div className="lbl">Quantity</div>
                  <input type="number" min="1" value={txQty} onChange={e=>setTxQty(e.target.value)} onBlur={e=>{const v=parseInt(e.target.value);if(!v||v<1)setTxQty("1");}} inputMode="numeric" style={{width:"100%"}}/>
                </div>
                <button className="btn btn-amber" onClick={handleTransaction} style={{fontSize:13,padding:"12px"}}>{txMode==="withdraw"?"⬆ Withdraw Keys":"⬇ Deposit Keys"}</button>
              </div>
            </div>
          )}

          {/* ADD PROPERTY */}
          {activeTab==="add"&&(
            <div style={{maxWidth:520}}>
              <div className="sec">Add New Property</div>
              <div className="fg">
                <div>
                  <div className="lbl">Lock Box</div>
                  <select value={addBox} onChange={e=>{setAddBox(e.target.value);setAddExisting(false);}} style={{width:"100%"}}>{boxes.map(b=><option key={b}>{b}</option>)}</select>
                  {boxStats[addBox]&&(
                    <div style={{fontSize:11,marginTop:4,color:pctColor(boxStats[addBox].pct),background:boxStats[addBox].pct>=100?"rgba(200,80,80,.08)":"transparent",padding:boxStats[addBox].pct>=100?"6px 10px":"0",borderRadius:4,border:boxStats[addBox].pct>=100?"1px solid rgba(200,80,80,.2)":"none"}}>
                      {boxStats[addBox].pct>=100?<>🔒 <strong>{addBox} FULL</strong> ({boxStats[addBox].props}/{boxStats[addBox].maxProps})</>:<>Capacity: {boxStats[addBox].props}/{boxStats[addBox].maxProps} ({boxStats[addBox].pct}% full)</>}
                    </div>
                  )}
                  {boxStats[addBox]&&boxStats[addBox].pct<100&&(()=>{
                    const boxRows=data[addBox]||[];
                    const totals=Object.fromEntries(KEY_TYPES.map(k=>[k,boxRows.reduce((s:number,r:PropertyRow)=>s+(r[k as KeyType]||0),0)]));
                    const grandTotal=KEY_TYPES.reduce((s:number,k:string)=>s+(totals[k]||0),0);
                    return (
                      <div style={{marginTop:8,background:"#0a0a08",border:"1px solid #2a2010",borderRadius:6,padding:"10px 14px"}}>
                        <div style={{fontSize:10,color:"#806040",letterSpacing:1,textTransform:"uppercase" as const,marginBottom:8}}>{addBox} - Current Keys</div>
                        <div style={{display:"flex",flexWrap:"wrap" as const,gap:6,marginBottom:8}}>
                          {KEY_TYPES.map(k=>(<div key={k} style={{display:"flex",alignItems:"center",gap:4,fontSize:12,borderRadius:6,padding:"3px 8px",color:(totals[k]||0)>0?"#c8960c":"#404040",background:(totals[k]||0)>0?"rgba(200,150,12,.08)":"rgba(50,50,50,.15)",border:`1px solid ${(totals[k]||0)>0?"rgba(200,150,12,.25)":"#252525"}`}}><span>{KEY_ICONS[k]}</span><span style={{fontSize:10,color:"#806040"}}>{k.replace(" Key","")}</span><strong style={{color:(totals[k]||0)>0?"#e8aa0e":"#505050"}}>{totals[k]||0}</strong></div>))}
                        </div>
                        <div style={{fontSize:11,color:"#504030",borderTop:"1px solid #1e1e10",paddingTop:6,display:"flex",justifyContent:"space-between"}}><span>Total</span><strong style={{color:"#c8960c"}}>{grandTotal}</strong></div>
                      </div>
                    );
                  })()}
                </div>
                <div>
                  <div className="lbl">Property Address</div>
                  <AddressInput value={addAddress}
                    onChange={v=>{setAddAddress(v);if(!v){setAddExisting(false);setAddKeys(Object.fromEntries(KEY_TYPES.map(k=>[k,"0"])));};}}
                    onSelect={(addr,box)=>{setAddAddress(addr);setAddBox(box);setAddKeys(Object.fromEntries(KEY_TYPES.map(k=>[k,"0"])));setAddExisting(true);}}
                    onAddNew={(q,box)=>{setAddAddress(q);setAddBox(box);setAddExisting(false);setAddKeys(Object.fromEntries(KEY_TYPES.map(k=>[k,"0"])));}}
                    allEntries={entries} boxes={boxes} boxStats={boxStats} placeholder="Enter property address..."/>
                  <div style={{fontSize:11,color:"#604030",marginTop:4}}>If address exists, keys will be <strong style={{color:"#c8960c"}}>added to existing</strong></div>
                </div>
                {addExisting&&addAddress&&(()=>{
                  const existRow=(data[addBox]||[]).find(r=>r["Property Address"].toLowerCase()===addAddress.toLowerCase());
                  if(!existRow) return null;
                  return (<div style={{background:"#0a0a08",border:"1px solid rgba(200,150,12,.3)",borderRadius:6,padding:"12px 14px"}}>
                    <div className="lbl" style={{color:"#c8960c",marginBottom:6}}>⚡ UPDATING - Lock Box #{existRow["Lock Box No."]}</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:8}}>{KEY_TYPES.map(k=>(<div key={k} style={{fontSize:12,color:(existRow[k]||0)>0?"#c8960c":"#404040"}}>{KEY_ICONS[k]} {k}: <strong>{existRow[k]||0}</strong></div>))}</div>
                  </div>);
                })()}
                <div>
                  <div className="lbl">Lock Box No.</div>
                  {(()=>{
                    const usedNos=new Set((data[addBox]||[]).map((r:PropertyRow)=>r["Lock Box No."]));
                    const maxNo=Math.max(...Array.from(usedNos),0);
                    const gapNos:number[]=[];for(let n=1;n<=maxNo;n++){if(!usedNos.has(n))gapNos.push(n);}
                    const nextNo=maxNo+1;
                    return (<div>
                      <div style={{display:"flex",gap:8,marginBottom:6}}>
                        <button className="btn" onClick={()=>setAddLockNo(String(nextNo))} style={{flex:1,padding:"7px 8px",fontSize:11,background:addLockNo===String(nextNo)?"#c8960c":"transparent",color:addLockNo===String(nextNo)?"#000":"#c8960c",border:"1px solid rgba(200,150,12,.4)",borderRadius:4}}>Next: #{nextNo}</button>
                        {gapNos.length>0&&<button className="btn" onClick={()=>setAddLockNo(String(gapNos[0]))} style={{flex:1,padding:"7px 8px",fontSize:11,background:addLockNo===String(gapNos[0])?"#50c880":"transparent",color:addLockNo===String(gapNos[0])?"#000":"#50c880",border:"1px solid rgba(80,200,128,.4)",borderRadius:4}}>Gap: #{gapNos[0]}</button>}
                      </div>
                      {gapNos.length>0&&(<div style={{marginBottom:6}}><div style={{fontSize:10,color:"#806040",marginBottom:4}}>AVAILABLE GAPS ({gapNos.length})</div><div style={{display:"flex",flexWrap:"wrap",gap:4,maxHeight:60,overflowY:"auto"}}>{gapNos.map(n=>(<div key={n} onClick={()=>setAddLockNo(String(n))} style={{padding:"3px 8px",borderRadius:4,cursor:"pointer",fontSize:12,fontWeight:"bold",background:addLockNo===String(n)?"#50c880":"rgba(80,200,128,.08)",color:addLockNo===String(n)?"#000":"#50c880",border:`1px solid ${addLockNo===String(n)?"#50c880":"rgba(80,200,128,.3)"}`}}>#{n}</div>))}</div></div>)}
                      <input type="number" value={addLockNo} onChange={e=>setAddLockNo(e.target.value)} onBlur={e=>{const v=parseInt(e.target.value);if(!v||v<1)setAddLockNo(String(nextNo));}} placeholder={`Auto (#${nextNo})`} inputMode="numeric" style={{width:"100%"}}/>
                    </div>);
                  })()}
                </div>
                <div style={{background:"#0a0a08",border:"1px solid #2a2010",borderRadius:6,padding:"14px"}}>
                  <div className="lbl" style={{marginBottom:12}}>{addExisting?"Keys to Add (summed)":"Initial Key Counts"}</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    {KEY_TYPES.map(k=>(<div key={k}><div style={{fontSize:11,color:"#806040",marginBottom:3}}>{KEY_ICONS[k]} {k}</div><input type="number" min="0" value={addKeys[k]} onChange={e=>setAddKeys(prev=>({...prev,[k]:e.target.value}))} onBlur={e=>{const v=parseInt(e.target.value);if(isNaN(v)||v<0)setAddKeys(prev=>({...prev,[k]:"0"}));}} inputMode="numeric" style={{width:"100%"}}/></div>))}
                  </div>
                </div>
                <button className="btn btn-amber" onClick={handleAddNew} style={{fontSize:13,padding:"12px",opacity:!addExisting&&(boxStats[addBox]?.pct||0)>=100?0.5:1,cursor:!addExisting&&(boxStats[addBox]?.pct||0)>=100?"not-allowed":"pointer"}}>
                  {addExisting?"➕ Add Keys to Existing":`➕ Add to ${addBox}`}
                </button>
                {addAnotherPrompt&&(<div style={{background:"rgba(80,200,120,.08)",border:"1px solid rgba(80,200,120,.25)",borderRadius:8,padding:"13px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
                  <div><div style={{fontSize:13,color:"#50c880",fontWeight:"bold"}}>✓ Added!</div><div style={{fontSize:11,color:"#806040",marginTop:2}}>Add another?</div></div>
                  <div style={{display:"flex",gap:8,flexShrink:0}}>
                    <button className="btn btn-amber" style={{padding:"7px 12px",fontSize:12}} onClick={()=>setAddAnotherPrompt(false)}>➕ Another</button>
                    <button className="btn btn-outline" style={{padding:"7px 12px",fontSize:12}} onClick={()=>{setAddAnotherPrompt(false);navigateTo("inventory");}}>View List</button>
                  </div>
                </div>)}
              </div>
            </div>
          )}

          {/* HISTORY */}
          {activeTab==="history"&&(
            <div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
                <div className="sec" style={{margin:0}}>History - {filteredLog.length} records</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <input placeholder="Search..." value={logSearch} onChange={e=>setLogSearch(e.target.value)} style={{width:200}}/>
                  <select value={logFilter} onChange={e=>setLogFilter(e.target.value)} style={{padding:"8px 10px"}}>
                    <option value="all">All</option><option value="deposit">Deposit</option><option value="withdraw">Withdraw</option><option value="add">New Added</option>
                  </select>
                  {txLog.length>0&&<button className="btn btn-outline" onClick={()=>{const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(txLog.map(l=>({Date:l.ts,Type:l.type.toUpperCase(),Address:l.address,"Lock Box":l.box,"Key Type":l.keyType,Qty:l.qty})));XLSX.utils.book_append_sheet(wb,ws,"History");XLSX.writeFile(wb,"RBR-Key_History.xlsx");showToast("Exported!");}}>⬇ Export</button>}
                </div>
              </div>
              <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
                {[{label:"Total",val:txLog.length,c:"#c8960c"},{label:"Deposits",val:txLog.filter(l=>l.type==="deposit").length,c:"#50c880"},{label:"Withdrawals",val:txLog.filter(l=>l.type==="withdraw").length,c:"#e07070"},{label:"New",val:txLog.filter(l=>l.type==="add").length,c:"#7090d0"}].map(s=>(
                  <div key={s.label} className="sc" style={{flex:1,minWidth:80}}><div style={{fontSize:18,fontWeight:"bold",color:s.c}}>{s.val}</div><div style={{fontSize:9,color:"#806040",letterSpacing:1,marginTop:2}}>{s.label.toUpperCase()}</div></div>
                ))}
              </div>
              {filteredLog.length===0?(
                <div style={{background:"#0c0c0a",border:"1px solid #1e1e10",borderRadius:8,padding:"50px 20px",textAlign:"center"}}>
                  <div style={{fontSize:30,marginBottom:10}}>📜</div><div style={{fontSize:13,color:"#604030"}}>No transactions yet</div>
                </div>
              ):(
                <div style={{background:"#0c0c0a",border:"1px solid #1e1e10",borderRadius:8,overflow:"hidden"}}>
                  <div style={{overflowX:"auto"}}>
                    <table>
                      <thead><tr><th>Date & Time</th><th>Type</th><th>Property Address</th><th>Box</th><th>Key Type</th><th>Qty</th></tr></thead>
                      <tbody>
                        {filteredLog.map(log=>(
                          <tr key={log.id}>
                            <td style={{color:"#806040",fontSize:11,whiteSpace:"nowrap"}}>{log.ts}</td>
                            <td>{log.type==="deposit"&&<span className="badge badge-dep">⬇ Deposit</span>}{log.type==="withdraw"&&<span className="badge badge-wit">⬆ Withdraw</span>}{log.type==="add"&&<span className="badge badge-add">➕ New</span>}</td>
                            <td style={{color:"#e0d0b0",maxWidth:240}}>{log.address}</td>
                            <td style={{color:"#c8960c",fontSize:12}}>{log.box}</td>
                            <td style={{fontSize:12,color:"#a09070"}}>{KEY_ICONS[log.keyType]||""} {log.keyType}</td>
                            <td style={{textAlign:"center"}}><span className="badge badge-a">{log.qty}</span></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* New Box Modal */}
      {showNewBox&&(<div className="mo" onClick={()=>setShowNewBox(false)}><div className="md" style={{width:380}} onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:13,color:"#c8960c",fontWeight:"bold",marginBottom:16,letterSpacing:1}}>📦 CREATE NEW BOX</div>
        <div className="fg">
          <div><div className="lbl">Box Name</div><input value={newBoxName} onChange={e=>setNewBoxName(e.target.value.toUpperCase())} placeholder="e.g. BOX 3" style={{width:"100%"}} autoFocus/></div>
          <div><div className="lbl">Max Properties</div><input type="number" min="1" value={newBoxMax} onChange={e=>setNewBoxMax(e.target.value)} onBlur={e=>{const v=parseInt(e.target.value);if(!v||v<1)setNewBoxMax("1");}} inputMode="numeric" style={{width:"100%"}}/></div>
          <div style={{display:"flex",gap:10}}>
            <button className="btn btn-amber" onClick={handleAddBox} style={{flex:1,padding:"11px"}}>Create</button>
            <button className="btn btn-outline" onClick={()=>setShowNewBox(false)} style={{flex:1,padding:"11px"}}>Cancel</button>
          </div>
        </div>
      </div></div>)}

      {/* Box Settings Modal */}
      {editBoxSettings&&(<div className="mo" onClick={()=>setEditBoxSettings(null)}><div className="md" style={{width:380}} onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:13,color:"#c8960c",fontWeight:"bold",marginBottom:6,letterSpacing:1}}>⚙ {editBoxSettings.name}</div>
        <div style={{fontSize:12,color:"#806040",marginBottom:16}}>Current: {boxStats[editBoxSettings.name]?.props||0} properties</div>
        <div className="fg">
          <div><div className="lbl">Max Properties</div><input type="number" min="1" value={editBoxSettings.maxProps} onChange={e=>setEditBoxSettings(s=>s?{...s,maxProps:e.target.value}:null)} onBlur={e=>{const v=parseInt(e.target.value);if(!v||v<1)setEditBoxSettings(s=>s?{...s,maxProps:"1"}:null);}} inputMode="numeric" style={{width:"100%"}}/><div style={{fontSize:11,color:"#604030",marginTop:4}}>Min: {boxStats[editBoxSettings.name]?.props||0}</div></div>
          <div style={{display:"flex",gap:10}}>
            <button className="btn btn-amber" onClick={handleSaveBoxSettings} style={{flex:1,padding:"11px"}}>Save</button>
            <button className="btn btn-outline" onClick={()=>setEditBoxSettings(null)} style={{flex:1,padding:"11px"}}>Cancel</button>
          </div>
        </div>
      </div></div>)}

      {/* Edit Row Modal */}
      {editRow&&(<div className="mo" onClick={()=>setEditRow(null)}><div className="md" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:13,color:"#c8960c",fontWeight:"bold",marginBottom:14,letterSpacing:1}}>✏ EDIT</div>
        <div className="fg">
          <div><div className="lbl">Property Address</div><input value={editRow["Property Address"]} onChange={e=>setEditRow(r=>r?{...r,"Property Address":e.target.value}:r)} style={{width:"100%"}}/></div>
          <div><div className="lbl">Lock Box No.</div><input type="number" value={editRow["Lock Box No."]} inputMode="numeric" onChange={e=>{const v=e.target.value;setEditRow(r=>r?{...r,"Lock Box No.":v===""?0:parseInt(v)||r["Lock Box No."]}:r);}} onBlur={e=>{if(!e.target.value)setEditRow(r=>r?{...r,"Lock Box No.":0}:r);}} style={{width:"100%"}}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            {KEY_TYPES.map(k=>(<div key={k}><div className="lbl">{KEY_ICONS[k]} {k}</div><input type="number" min="0" value={editRow[k]} inputMode="numeric" onChange={e=>{const v=e.target.value;setEditRow(r=>r?{...r,[k]:v===""?0:parseInt(v)||0}:r);}} onBlur={e=>{if(e.target.value===""||parseInt(e.target.value)<0)setEditRow(r=>r?{...r,[k]:0}:r);}} style={{width:"100%"}}/></div>))}
          </div>
          <div style={{display:"flex",gap:10}}>
            <button className="btn btn-amber" onClick={handleSaveEdit} style={{flex:1,padding:"11px"}}>💾 Save</button>
            <button className="btn btn-outline" onClick={()=>setEditRow(null)} style={{flex:1,padding:"11px"}}>Cancel</button>
          </div>
          <button className="btn btn-danger" style={{width:"100%",padding:"10px",fontSize:12}} onClick={()=>{if(editRow)handleDelete(editRow._box,editRow._id);}}>🗑 Delete This Property</button>
        </div>
      </div></div>)}

      {/* Confirm Delete */}
      {confirmDelete&&(<div className="mo" onClick={()=>setConfirmDelete(null)}><div className="md" style={{width:360,textAlign:"center"}} onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:32,marginBottom:10}}>🗑</div>
        <div style={{fontSize:14,color:"#e8e0d0",fontWeight:"bold",marginBottom:8}}>Delete Property?</div>
        <div style={{fontSize:12,color:"#c05050",background:"rgba(200,80,80,.08)",border:"1px solid rgba(200,80,80,.2)",borderRadius:6,padding:"10px 14px",marginBottom:18}}>"{confirmDelete.address}"<div style={{fontSize:11,color:"#806040",marginTop:4}}>This cannot be undone.</div></div>
        <div style={{display:"flex",gap:10}}>
          <button className="btn btn-outline" onClick={()=>setConfirmDelete(null)} style={{flex:1,padding:"12px"}}>Cancel</button>
          <button onClick={confirmDeleteNow} style={{flex:1,padding:"12px",background:"#c05050",color:"#fff",border:"none",borderRadius:4,cursor:"pointer",fontFamily:'"Courier New",monospace',fontWeight:"bold",fontSize:13}}>🗑 Yes, Delete</button>
        </div>
      </div></div>)}

      {/* Clear Data Modal */}
      {showClearData&&(<div className="mo" onClick={()=>{setShowClearData(false);setClearStep(1);setClearConfirmText("");setClearPinDigits([]);setClearPinError(false);}}>
        <div className="md" style={{width:400}} onClick={e=>e.stopPropagation()}>
          <div style={{fontSize:26,textAlign:"center",marginBottom:8}}>⚠️</div>
          <div style={{fontSize:14,color:"#e06060",fontWeight:"bold",textAlign:"center",marginBottom:4,letterSpacing:1}}>CLEAR DATA</div>
          <div style={{display:"flex",justifyContent:"center",gap:6,marginBottom:18}}>
            {[1,2,3].map(s=>(<div key={s} style={{width:28,height:4,borderRadius:2,background:clearStep>=s?"#c05050":"#2a2010",transition:"background .2s"}}/>))}
          </div>

          {clearStep===1&&(
            <>
              <div className="lbl" style={{marginBottom:10}}>Select what to delete</div>
              {boxes.map(box=>(
                <div key={box} onClick={()=>setClearTarget(box)} style={{padding:"10px 14px",borderRadius:6,cursor:"pointer",marginBottom:6,border:`1px solid ${clearTarget===box?"#c05050":"#2a2010"}`,background:clearTarget===box?"rgba(200,80,80,.1)":"transparent",display:"flex",alignItems:"center",justifyContent:"space-between",transition:"all .15s"}}>
                  <div><div style={{fontSize:13,color:clearTarget===box?"#e06060":"#e0d0b0",fontWeight:"bold"}}>📦 {box}</div><div style={{fontSize:10,color:"#604030",marginTop:2}}>{(data[box]||[]).length} properties · {boxStats[box]?.keys||0} keys</div></div>
                  <div style={{fontSize:11,color:clearTarget===box?"#e06060":"#504030",border:`1px solid ${clearTarget===box?"rgba(200,80,80,.4)":"#2a2010"}`,padding:"2px 8px",borderRadius:4}}>{clearTarget===box?"✓":"Select"}</div>
                </div>
              ))}
              <div onClick={()=>setClearTarget("history")} style={{padding:"10px 14px",borderRadius:6,cursor:"pointer",marginBottom:6,border:`1px solid ${clearTarget==="history"?"#c05050":"#2a2010"}`,background:clearTarget==="history"?"rgba(200,80,80,.1)":"transparent",display:"flex",alignItems:"center",justifyContent:"space-between",transition:"all .15s"}}>
                <div><div style={{fontSize:13,color:clearTarget==="history"?"#e06060":"#e0d0b0",fontWeight:"bold"}}>📜 History Only</div><div style={{fontSize:10,color:"#604030",marginTop:2}}>{txLog.length} records</div></div>
                <div style={{fontSize:11,color:clearTarget==="history"?"#e06060":"#504030",border:`1px solid ${clearTarget==="history"?"rgba(200,80,80,.4)":"#2a2010"}`,padding:"2px 8px",borderRadius:4}}>{clearTarget==="history"?"✓":"Select"}</div>
              </div>
              <div onClick={()=>setClearTarget("all")} style={{padding:"10px 14px",borderRadius:6,cursor:"pointer",marginBottom:18,border:`1px solid ${clearTarget==="all"?"#c05050":"#3a1010"}`,background:clearTarget==="all"?"rgba(200,80,80,.15)":"rgba(60,10,10,.2)",display:"flex",alignItems:"center",justifyContent:"space-between",transition:"all .15s"}}>
                <div><div style={{fontSize:13,color:"#e06060",fontWeight:"bold"}}>🗑 Everything</div><div style={{fontSize:10,color:"#604030",marginTop:2}}>{Object.values(data).flat().length} properties + {txLog.length} records</div></div>
                <div style={{fontSize:11,color:clearTarget==="all"?"#e06060":"#504030",border:`1px solid ${clearTarget==="all"?"rgba(200,80,80,.4)":"#3a1010"}`,padding:"2px 8px",borderRadius:4}}>{clearTarget==="all"?"✓":"Select"}</div>
              </div>
              <div style={{display:"flex",gap:10}}>
                <button className="btn btn-outline" onClick={()=>setShowClearData(false)} style={{flex:1,padding:"11px"}}>Cancel</button>
                <button onClick={()=>setClearStep(2)} style={{flex:1,padding:"11px",background:"#c05050",color:"#fff",border:"none",borderRadius:4,cursor:"pointer",fontFamily:'"Courier New",monospace',fontWeight:"bold",fontSize:13}}>Next →</button>
              </div>
            </>
          )}

          {clearStep===2&&(
            <>
              <div style={{background:"rgba(200,80,80,.08)",border:"1px solid rgba(200,80,80,.2)",borderRadius:6,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#e06060"}}>
                Deleting: <strong>{clearTarget==="all"?"ALL DATA":clearTarget==="history"?"History Only":clearTarget}</strong>
              </div>
              <div style={{fontSize:13,color:"#806040",textAlign:"center",marginBottom:14}}>Enter PIN to authorize</div>
              <div className={clearPinShake?"shake":""} style={{display:"flex",justifyContent:"center",gap:12,marginBottom:18}}>
                {[0,1,2,3].map(i=>(<div key={i} style={{width:14,height:14,borderRadius:"50%",background:i<clearPinDigits.length?(clearPinError?"#e06060":"#c8960c"):"transparent",border:`2px solid ${i<clearPinDigits.length?(clearPinError?"#e06060":"#c8960c"):"#3a3020"}`,transition:"all .15s"}}/>))}
              </div>
              {clearPinError&&<div style={{fontSize:11,color:"#e06060",textAlign:"center",marginBottom:10}}>Incorrect PIN</div>}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
                {["1","2","3","4","5","6","7","8","9","","0","⌫"].map((k,i)=>(
                  k===""?<div key={i}/>:
                  k==="⌫"?<div key={i} className="pin-key" onClick={()=>setClearPinDigits(p=>p.slice(0,-1))} style={{fontSize:18,color:"#806040",height:56,borderRadius:8}}>⌫</div>:
                  <div key={i} className="pin-key" onClick={()=>handleClearPinDigit(k)} style={{fontSize:20,height:56,borderRadius:8}}>{k}</div>
                ))}
              </div>
              <button className="btn btn-outline" onClick={()=>{setClearStep(1);setClearPinDigits([]);setClearPinError(false);}} style={{width:"100%",padding:"10px"}}>← Back</button>
            </>
          )}

          {clearStep===3&&(
            <>
              <div style={{background:"rgba(200,80,80,.08)",border:"1px solid rgba(200,80,80,.25)",borderRadius:6,padding:"12px 14px",marginBottom:14}}>
                <div style={{fontSize:12,color:"#e06060",fontWeight:"bold",marginBottom:4}}>⚠ Deleting: {clearTarget==="all"?"ALL DATA":clearTarget==="history"?"History":clearTarget}</div>
                <div style={{fontSize:11,color:"#806040"}}>{clearTarget==="all"?`${Object.values(data).flat().length} properties + ${txLog.length} records`:clearTarget==="history"?`${txLog.length} records`:`${(data[clearTarget]||[]).length} properties`} will be permanently deleted</div>
              </div>
              <div className="lbl" style={{marginBottom:6}}>Type <strong style={{color:"#e06060",letterSpacing:2}}>DELETE</strong> to confirm</div>
              <input value={clearConfirmText} onChange={e=>setClearConfirmText(e.target.value)} placeholder='Type DELETE here' autoFocus style={{width:"100%",marginBottom:14,borderColor:clearConfirmText==="DELETE"?"#c05050":undefined}}/>
              <div style={{display:"flex",gap:10}}>
                <button className="btn btn-outline" onClick={()=>{setClearStep(2);setClearConfirmText("");}} style={{flex:1,padding:"11px"}}>← Back</button>
                <button onClick={handleClearData} disabled={clearConfirmText!=="DELETE"} style={{flex:1,padding:"11px",background:clearConfirmText==="DELETE"?"#c05050":"#2a1010",color:clearConfirmText==="DELETE"?"#fff":"#604040",border:"none",borderRadius:4,cursor:clearConfirmText==="DELETE"?"pointer":"not-allowed",fontFamily:'"Courier New",monospace',fontWeight:"bold",fontSize:13,transition:"all .2s"}}>🗑 Delete Now</button>
              </div>
            </>
          )}
        </div>
      </div>)}

      {/* Toast */}
      {toast&&(<div style={{position:"fixed",bottom:22,right:22,zIndex:2000,background:toast.type==="error"?"#3a0a0a":"#0a1a0a",border:`1px solid ${toast.type==="error"?"#8a2020":"#207020"}`,color:toast.type==="error"?"#ff8080":"#80d080",padding:"11px 20px",borderRadius:8,fontSize:13,fontFamily:'"Courier New",monospace',boxShadow:"0 4px 20px rgba(0,0,0,.5)",animation:"slideIn .2s ease",maxWidth:400}}>{toast.msg}</div>)}
    </div>
  );
}

