import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";

const STUDIOS = [
  { id:"스튜디오(지하)", color:"#1D9E75" },
  { id:"17층1호",        color:"#378ADD" },
  { id:"17층2호",        color:"#7F77DD" },
  { id:"17층3호",        color:"#D4537E" },
  { id:"온택트룸2",      color:"#D85A30" },
  { id:"온택트룸4",      color:"#EF9F27" },
  { id:"온택트룸6",      color:"#639922" },
  { id:"대전",           color:"#888780" },
  { id:"광주",           color:"#0C447C" },
  { id:"구미",           color:"#993C1D" },
];

const ALIASES = {
  "지하":"스튜디오(지하)","스튜디오지하":"스튜디오(지하)","스튜디오(지하)":"스튜디오(지하)",
  "17층1호":"17층1호","17층 1호":"17층1호",
  "17층2호":"17층2호","17층 2호":"17층2호",
  "17층3호":"17층3호","17층 3호":"17층3호",
  "온택트룸2":"온택트룸2","온택트룸 2":"온택트룸2","온택2":"온택트룸2",
  "온택트룸4":"온택트룸4","온택트룸 4":"온택트룸4","온택4":"온택트룸4",
  "온택트룸6":"온택트룸6","온택트룸 6":"온택트룸6","온택6":"온택트룸6",
  "대전":"대전","광주":"광주","구미":"구미",
};

const DAYS = ["월","화","수","목","금"];
const HOURS = Array.from({length:12},(_,i)=>`${String(i+7).padStart(2,"0")}:00`);

function excelDateToStr(val){
  if(!val&&val!==0) return "";
  const s=String(val).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if(typeof val==="number"&&val>1000){
    const d=new Date(Math.round((val-25569)*86400*1000));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
  }
  return s;
}
function excelTimeToStr(val){
  if(!val&&val!==0) return "";
  const s=String(val).trim();
  if(/^\d{1,2}:\d{2}$/.test(s)) return s.padStart(5,"0");
  if(typeof val==="number"&&val>=0&&val<1){
    const m=Math.round(val*24*60);
    return `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;
  }
  return s;
}
function calcLen(s,e){
  const sm=toMin(s),em=toMin(e);
  if(sm===null||em===null||em<=sm) return "";
  const d=em-sm;
  return `${String(Math.floor(d/60)).padStart(2,"0")}:${String(d%60).padStart(2,"0")}`;
}
function normStudio(raw){
  if(!raw) return "";
  const t=String(raw).trim();
  return ALIASES[t]||ALIASES[t.replace(/\s/g,"")]||t;
}
function toMin(t){
  if(!t) return null;
  const p=String(t).split(":");
  if(p.length<2) return null;
  const h=parseInt(p[0]),m=parseInt(p[1]);
  return isNaN(h)||isNaN(m)?null:h*60+m;
}
function getThisWeekMonday(){
  const d=new Date();const dow=d.getDay()||7;
  d.setDate(d.getDate()-dow+1);return d;
}
function addDays(d,n){const r=new Date(d);r.setDate(r.getDate()+n);return r;}
function fmtShort(d){return `${d.getMonth()+1}/${d.getDate()}`;}
function fmtFull(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
function getColor(id){return STUDIOS.find(s=>s.id===id)?.color||"#888";}

function detectConflicts(rows){
  const res=[];
  for(let i=0;i<rows.length;i++) for(let j=i+1;j<rows.length;j++){
    const a=rows[i],b=rows[j];
    if(!a.장소||a.장소!==b.장소||a.날짜!==b.날짜) continue;
    const aS=toMin(a.시작시간),aE=toMin(a.종료시간),bS=toMin(b.시작시간),bE=toMin(b.종료시간);
    if(aS!==null&&aE!==null&&bS!==null&&bE!==null&&aS<bE&&bS<aE)
      res.push({a:i,b:j,studio:a.장소,date:a.날짜,timeA:`${a.시작시간}~${a.종료시간}`,timeB:`${b.시작시간}~${b.종료시간}`});
  }
  return res;
}

const bd=(c)=>{const m={red:{bg:"#FCEBEB",fg:"#A32D2D"},green:{bg:"#E1F5EE",fg:"#0F6E56"},blue:{bg:"#E6F1FB",fg:"#185FA5"},amber:{bg:"#FAEEDA",fg:"#854F0B"},gray:{bg:"#f0f0ee",fg:"#666"}};const x=m[c]||m.gray;return{display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:500,background:x.bg,color:x.fg};};
const btn={display:"inline-flex",alignItems:"center",gap:5,padding:"0 14px",height:34,borderRadius:8,border:"0.5px solid #ccc",fontSize:13,cursor:"pointer",background:"#fff",color:"#333",fontFamily:"inherit"};
const btnP={...btn,background:"#1D9E75",borderColor:"#1D9E75",color:"#fff"};
const btnR={...btn,background:"#FCEBEB",borderColor:"#F09595",color:"#A32D2D"};
const panel={background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,padding:"16px 18px",marginBottom:12};
const inp={width:"100%",height:36,padding:"0 10px",borderRadius:8,border:"0.5px solid #ccc",fontSize:13,fontFamily:"inherit",background:"#fff",color:"#1c1c1a",boxSizing:"border-box"};
const lbl={fontSize:12,fontWeight:500,color:"#555",marginBottom:4,display:"block"};

function BookingForm({initial,onSave,onClose,title}){
  const [form,setForm]=useState(initial||{장소:"",날짜:"",구분:"강의",주제:"",내용:"",강사명:"",시작시간:"09:00",종료시간:"18:00",요일:""});
  const [err,setErr]=useState("");
  function set(k,v){let u={...form,[k]:v};if(k==="날짜"&&v){const d=new Date(v);u.요일=["일","월","화","수","목","금","토"][d.getDay()];}setForm(u);}
  const len=calcLen(form.시작시간,form.종료시간);
  function save(){
    if(!form.장소) return setErr("스튜디오를 선택해주세요");
    if(!form.날짜) return setErr("날짜를 입력해주세요");
    if(toMin(form.시작시간)>=toMin(form.종료시간)) return setErr("종료 시간이 시작 시간보다 늦어야 합니다");
    onSave({...form,길이:len}); onClose();
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"#fff",borderRadius:16,width:520,maxHeight:"92vh",overflowY:"auto",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"18px 22px",borderBottom:"0.5px solid #e5e5e3",position:"sticky",top:0,background:"#fff"}}>
          <div style={{fontSize:16,fontWeight:500}}>{title}</div>
          <button style={{...btn,height:28,padding:"0 10px"}} onClick={onClose}>✕</button>
        </div>
        <div style={{padding:"18px 22px",display:"flex",flexDirection:"column",gap:14}}>
          <div>
            <span style={lbl}>스튜디오 *</span>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6}}>
              {STUDIOS.map(s=>(
                <label key={s.id} style={{display:"flex",alignItems:"center",gap:7,padding:"7px 10px",border:`1.5px solid ${form.장소===s.id?s.color:"#e5e5e3"}`,borderRadius:8,cursor:"pointer",background:form.장소===s.id?"#f8fffe":"#fff",fontSize:12}}>
                  <input type="radio" name="bst" style={{display:"none"}} checked={form.장소===s.id} onChange={()=>set("장소",s.id)}/>
                  <span style={{width:9,height:9,borderRadius:"50%",background:s.color,flexShrink:0}}></span>
                  {s.id}{form.장소===s.id&&<span style={{marginLeft:"auto",color:s.color,fontWeight:600}}>✓</span>}
                </label>
              ))}
            </div>
          </div>
          <div>
            <span style={lbl}>구분 *</span>
            <div style={{display:"flex",gap:6}}>
              {["방송준비","강의"].map(g=>(
                <label key={g} style={{display:"flex",alignItems:"center",gap:6,padding:"7px 18px",border:`1.5px solid ${form.구분===g?"#378ADD":"#e5e5e3"}`,borderRadius:8,cursor:"pointer",fontSize:12,background:form.구분===g?"#F0F6FF":"#fff"}}>
                  <input type="radio" name="bgb" style={{display:"none"}} checked={form.구분===g} onChange={()=>set("구분",g)}/>{g}
                </label>
              ))}
            </div>
          </div>
          <div>
            <span style={lbl}>날짜 *</span>
            <input type="date" style={inp} value={form.날짜} onChange={e=>set("날짜",e.target.value)}/>
            {form.요일&&<div style={{fontSize:11,color:"#1D9E75",marginTop:3}}>{form.요일}요일</div>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><span style={lbl}>시작 시간 *</span><input type="time" style={inp} value={form.시작시간} onChange={e=>set("시작시간",e.target.value)}/></div>
            <div><span style={lbl}>종료 시간 *</span><input type="time" style={inp} value={form.종료시간} onChange={e=>set("종료시간",e.target.value)}/></div>
          </div>
          {len&&<div style={{fontSize:12,color:"#1D9E75",fontWeight:500}}>⏱ 총 {len}</div>}
          <div><span style={lbl}>주제</span><input type="text" style={inp} placeholder="예: 코딩 Live Django트랙" value={form.주제} onChange={e=>set("주제",e.target.value)}/></div>
          <div><span style={lbl}>내용</span><input type="text" style={inp} placeholder="예: Django : DRF 1" value={form.내용} onChange={e=>set("내용",e.target.value)}/></div>
          <div><span style={lbl}>강사명</span><input type="text" style={inp} placeholder="예: 김준호 강사" value={form.강사명} onChange={e=>set("강사명",e.target.value)}/></div>
          {err&&<div style={{fontSize:12,color:"#E24B4A",background:"#FFF0F0",padding:"8px 12px",borderRadius:8}}>{err}</div>}
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",paddingTop:4,borderTop:"0.5px solid #f0f0ee"}}>
            <button style={btn} onClick={onClose}>취소</button>
            <button style={btnP} onClick={save}>{title==="예약 수정"?"수정 저장":"예약 등록"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function CancelModal({row,onClose,onConfirm}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}}>
      <div style={{background:"#fff",borderRadius:16,width:390,padding:"24px",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{fontSize:16,fontWeight:500,marginBottom:12}}>예약을 취소하시겠습니까?</div>
        <div style={{fontSize:13,color:"#555",marginBottom:12,lineHeight:1.9,background:"#f8f8f6",borderRadius:8,padding:"10px 14px"}}>
          <div><b>스튜디오:</b> {row?.장소}</div>
          <div><b>날짜:</b> {row?.날짜} {row?.요일&&`(${row.요일}요일)`}</div>
          <div><b>시간:</b> {row?.시작시간} ~ {row?.종료시간}</div>
          <div><b>내용:</b> {row?.내용||row?.주제||"-"}</div>
        </div>
        <div style={{background:"#FFF8F8",border:"0.5px solid #F09595",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#A32D2D",marginBottom:16}}>⚠ 취소된 예약은 복구할 수 없습니다.</div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button style={btn} onClick={onClose}>돌아가기</button>
          <button style={btnR} onClick={onConfirm}>예약 취소 확인</button>
        </div>
      </div>
    </div>
  );
}

function BlockPopup({row,idx,pos,onClose,onEdit,onCancel}){
  const color=getColor(row.장소);
  return(
    <div style={{position:"fixed",top:Math.min(pos.y+10,window.innerHeight-320),left:Math.min(pos.x+10,window.innerWidth-280),zIndex:3000,background:"#fff",borderRadius:12,border:"0.5px solid #e5e5e3",boxShadow:"0 8px 28px rgba(0,0,0,0.18)",width:260,padding:"14px 16px"}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{display:"flex",alignItems:"center",gap:6,fontWeight:500,fontSize:13,color}}>
          <span style={{width:9,height:9,borderRadius:"50%",background:color}}></span>{row.장소}
        </span>
        <button style={{background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}} onClick={onClose}>✕</button>
      </div>
      <div style={{fontSize:12,lineHeight:1.9,color:"#444",borderBottom:"0.5px solid #f0f0ee",paddingBottom:10,marginBottom:10}}>
        {row._conflict&&<div style={{color:"#E24B4A",fontWeight:600,marginBottom:4}}>⚠ 충돌</div>}
        <div><b>구분:</b> {row.구분}</div>
        <div><b>내용:</b> {row.내용||row.주제||"-"}</div>
        {row.강사명&&<div><b>강사:</b> {row.강사명}</div>}
        <div><b>날짜:</b> {row.날짜} {row.요일&&`(${row.요일})`}</div>
        <div><b>시간:</b> {row.시작시간} ~ {row.종료시간}</div>
      </div>
      <div style={{display:"flex",gap:6}}>
        <button style={{...btn,flex:1,height:30,fontSize:12,justifyContent:"center"}} onClick={()=>{onEdit(idx);onClose();}}>✏ 수정</button>
        <button style={{...btnR,flex:1,height:30,fontSize:12,justifyContent:"center"}} onClick={()=>{onCancel(idx);onClose();}}>✕ 취소</button>
      </div>
    </div>
  );
}

function WeeklyGrid({rows,conflicts,monday,onEdit,onCancel}){
  const CELL_H=60,TIME_W=52;
  const [popup,setPopup]=useState(null);
  const gridRef=useRef();
  const dayDates=DAYS.map((_,i)=>fmtFull(addDays(monday,i)));
  const conflictIdxs=new Set();
  conflicts.forEach(c=>{conflictIdxs.add(c.a);conflictIdxs.add(c.b);});

  useEffect(()=>{
    function measure(){if(gridRef.current){const w=gridRef.current.getBoundingClientRect().width;}}
    measure();
    window.addEventListener("resize",measure);
    return()=>window.removeEventListener("resize",measure);
  },[]);

  function layoutBlocks(blocks){
    const sorted=[...blocks].sort((a,b)=>toMin(a.시작시간)-toMin(b.시작시간));
    const cols=[];
    const result=sorted.map(b=>{
      const sMin=toMin(b.시작시간),eMin=toMin(b.종료시간);
      let col=0;while(cols[col]&&cols[col]>sMin)col++;
      cols[col]=eMin;return{...b,col,totalCols:1};
    });
    result.forEach(b=>{
      const sMin=toMin(b.시작시간),eMin=toMin(b.종료시간);
      const ov=result.filter(o=>toMin(o.시작시간)<eMin&&toMin(o.종료시간)>sMin);
      const mx=Math.max(...ov.map(o=>o.col));
      ov.forEach(o=>{o.totalCols=Math.max(o.totalCols,mx+1);});
    });
    return result;
  }

  return(
    <div ref={gridRef} style={{position:"relative",width:"100%"}} onClick={()=>setPopup(null)}>
      {popup&&<BlockPopup row={popup.row} idx={popup.idx} pos={{x:popup.x,y:popup.y}} onClose={()=>setPopup(null)} onEdit={onEdit} onCancel={onCancel}/>}
      <div style={{display:"flex",background:"#fff",width:"100%",overflow:"hidden"}}>
        <div style={{width:TIME_W,flexShrink:0,borderRight:"0.5px solid #e5e5e3",background:"#fafaf9"}}>
          <div style={{height:52,borderBottom:"0.5px solid #e5e5e3",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#bbb",background:"#f8f8f6"}}>시간</div>
          {HOURS.map(h=><div key={h} style={{height:CELL_H,borderBottom:"0.5px solid #f0f0ee",display:"flex",alignItems:"flex-start",justifyContent:"center",paddingTop:4,fontSize:11,color:"#ccc",fontWeight:500}}>{h}</div>)}
        </div>
        <div style={{flex:1,display:"flex"}}>
          {DAYS.map((day,di)=>{
            const dayDate=dayDates[di];
            const isToday=fmtFull(new Date())===dayDate;
            const blocks=layoutBlocks(rows.map((r,i)=>({...r,idx:i})).filter(r=>r.날짜===dayDate&&r.장소&&r.시작시간&&r.종료시간));
            return(
              <div key={day} style={{flex:1,borderRight:"0.5px solid #e5e5e3",position:"relative",minWidth:0}}>
                <div style={{height:52,borderBottom:"0.5px solid #e5e5e3",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",background:isToday?"#F0FBF6":"#f8f8f6"}}>
                  <span style={{fontSize:14,fontWeight:600,color:isToday?"#1D9E75":"#222"}}>{day}요일</span>
                  <span style={{fontSize:12,color:isToday?"#1D9E75":"#aaa"}}>{fmtShort(addDays(monday,di))}{isToday&&" · 오늘"}</span>
                </div>
                <div style={{position:"relative"}}>
                  {HOURS.map(h=><div key={h} style={{height:CELL_H,borderBottom:"0.5px solid #f0f0ee",background:isToday?"#FAFFFE":"#fff"}}/>)}
                  {blocks.length===0&&<div style={{position:"absolute",top:"45%",width:"100%",textAlign:"center",fontSize:12,color:"#e0e0de",pointerEvents:"none"}}>예약 없음</div>}
                  {blocks.map((b,bi)=>{
                    const sMin=toMin(b.시작시간),eMin=toMin(b.종료시간);
                    if(sMin===null||eMin===null) return null;
                    const top=(sMin-7*60)/60*CELL_H+1;
                    const height=Math.max((eMin-sMin)/60*CELL_H-3,24);
                    const isCf=conflictIdxs.has(b.idx);
                    const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                    const color=isCf?"#E24B4A":getColor(b.장소);
                    const bg=isCf?"rgba(226,74,74,0.07)":isPrep?"rgba(29,158,117,0.07)":"rgba(55,138,221,0.07)";
                    const colW=100/b.totalCols;
                    const content=b.내용||b.주제||"";
                    const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                    return(
                      <div key={bi}
                        style={{position:"absolute",top,left:`${b.col*colW+0.5}%`,width:`${colW-1}%`,height,background:bg,border:`1.5px solid ${color}`,borderLeft:`3px solid ${color}`,borderRadius:6,padding:"4px 7px",overflow:"hidden",boxSizing:"border-box",cursor:"pointer"}}
                        onClick={e=>{e.stopPropagation();setPopup({row:{...b,_conflict:conflictIdxs.has(b.idx)},idx:b.idx,x:e.clientX,y:e.clientY});}}
                      >
                        {isCf&&<div style={{fontSize:10,color:"#E24B4A",fontWeight:700,lineHeight:1.3}}>⚠ 충돌</div>}
                        <div style={{fontSize:11,fontWeight:700,color,lineHeight:1.4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{b.장소}</div>
                        {height>36&&<div style={{fontSize:11,color:"#333",lineHeight:1.4,overflow:"hidden",textOverflow:"ellipsis"}}>{isPrep?"방송준비":shortContent}</div>}
                        {height>58&&b.강사명&&<div style={{fontSize:10,color:"#888",lineHeight:1.3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{b.강사명}</div>}
                        {height>74&&<div style={{fontSize:10,color:"#aaa",lineHeight:1.3}}>{b.시작시간}~{b.종료시간}</div>}
                        {height>90&&(
                          <div style={{display:"flex",gap:4,marginTop:4}}>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:`0.5px solid ${color}`,borderRadius:4,fontSize:10,cursor:"pointer",padding:"2px 0",color,fontWeight:600}} onClick={e=>{e.stopPropagation();onEdit(b.idx);}}>✏</button>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:"0.5px solid #F09595",borderRadius:4,fontSize:10,cursor:"pointer",padding:"2px 0",color:"#A32D2D",fontWeight:600}} onClick={e=>{e.stopPropagation();onCancel(b.idx);}}>✕</button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

export default function App(){
  const [tab,setTab]=useState("dashboard");
  const [rows,setRows]=useState([]);
  const [conflicts,setConflicts]=useState([]);
  const [notifs,setNotifs]=useState([{id:0,type:"info",title:"시스템 준비 완료",desc:"편성표 업로드 또는 직접 예약 등록 후 대시보드에서 확인하세요.",time:"",unread:false}]);
  const [uploadState,setUploadState]=useState("idle");
  const [parseLog,setParseLog]=useState([]);
  const [weekOffset,setWeekOffset]=useState(0);
  const [bookingModal,setBookingModal]=useState(null);
  const [cancelTarget,setCancelTarget]=useState(null);
  const [filterStudio,setFilterStudio]=useState("전체");
  const [fileName,setFileName]=useState("");
  const [saving,setSaving]=useState(false);
  const fileRef=useRef();

  // Supabase에서 데이터 로드
  useEffect(()=>{
    async function load(){
      const {data,error}=await supabase
        .from("bookings")
        .select("*")
        .order("created_at",{ascending:true});
      if(!error&&data){
        const mapped=data.map(r=>({
          _id:r.id,
          구분:r.구분||"",장소:r.장소||"",주제:r.주제||"",
          내용:r.내용||"",강사명:r.강사명||"",날짜:r.날짜||"",
          요일:r.요일||"",시작시간:r.시작시간||"",종료시간:r.종료시간||"",
          길이:r.길이||"",_src:r.src||"manual",
        }));
        setRows(mapped);
        setConflicts(detectConflicts(mapped));
      }
    }
    load();
  },[]);

  async function reCalcAndSave(newRows){
    setRows(newRows);
    setConflicts(detectConflicts(newRows));
  }

  async function dbInsert(row){
    setSaving(true);
    const {data,error}=await supabase.from("bookings").insert([{
      구분:row.구분,장소:row.장소,주제:row.주제,내용:row.내용,
      강사명:row.강사명,날짜:row.날짜,요일:row.요일,
      시작시간:row.시작시간,종료시간:row.종료시간,길이:row.길이,src:row._src||"manual",
    }]).select();
    setSaving(false);
    return data?.[0];
  }

  async function dbUpdate(id,row){
    setSaving(true);
    await supabase.from("bookings").update({
      구분:row.구분,장소:row.장소,주제:row.주제,내용:row.내용,
      강사명:row.강사명,날짜:row.날짜,요일:row.요일,
      시작시간:row.시작시간,종료시간:row.종료시간,길이:row.길이,
    }).eq("id",id);
    setSaving(false);
  }

  async function dbDelete(id){
    await supabase.from("bookings").delete().eq("id",id);
  }

  function handleFile(e){
    const file=e.target.files[0];
    if(!file) return;
    setFileName(file.name); setUploadState("parsing");
    const reader=new FileReader();
    reader.onload=async(evt)=>{
      try{
        const wb=XLSX.read(evt.target.result,{type:"binary",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
        const parsed=raw.map(r=>{
          const 날짜=excelDateToStr(r["날짜"]||r["F"]||"");
          const 시작=excelTimeToStr(r["시작시간"]||r["H"]||"");
          const 종료=excelTimeToStr(r["종료시간"]||r["I"]||"");
          return{구분:String(r["구분"]||r["A"]||""),장소:normStudio(r["장소"]||r["B"]||""),주제:String(r["주제"]||r["C"]||""),내용:String(r["내용"]||r["D"]||""),강사명:String(r["강사명"]||r["E"]||""),날짜,요일:String(r["요일"]||r["G"]||""),시작시간:시작,종료시간:종료,길이:calcLen(시작,종료),_src:"excel"};
        }).filter(r=>(r.구분||r.주제||r.내용)&&r.날짜);

        // Supabase에 일괄 저장
        setSaving(true);
        const {data}=await supabase.from("bookings").insert(
          parsed.map(r=>({구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이,src:"excel"}))
        ).select();
        setSaving(false);

        const saved=(data||[]).map(r=>({_id:r.id,구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이,_src:"excel"}));
        const merged=[...rows,...saved];
        const cfls=detectConflicts(merged);
        reCalcAndSave(merged);
        setUploadState("done");
        const studios=[...new Set(saved.map(r=>r.장소).filter(Boolean))];
        setParseLog([`✅ ${saved.length}행 저장 완료`,`🏢 감지된 스튜디오: ${studios.join(", ")||"없음"}`,cfls.length>0?`⚠️ 충돌 ${cfls.length}건`:`✓ 충돌 없음`]);
        const newN=[];
        cfls.forEach(cf=>newN.push({id:Date.now()+Math.random(),type:"conflict",title:`⚠ 충돌 — ${cf.studio} · ${cf.date}`,desc:`${cf.timeA} ↔ ${cf.timeB}`,time:"방금",unread:true,rowA:cf.a,rowB:cf.b}));
        newN.push({id:Date.now(),type:"upload",title:`업로드 완료 — ${file.name}`,desc:`${studios.join(", ")} ${saved.length}건 저장됨`,time:"방금",unread:true});
        setNotifs(prev=>[...newN,...prev]);
      }catch(err){setUploadState("error");setParseLog(["❌ 파싱 실패: "+err.message]);}
    };
    reader.readAsBinaryString(file);
  }

  async function handleSave(form,editIdx){
    if(editIdx!=null){
      const target=rows[editIdx];
      await dbUpdate(target._id,form);
      const newRows=rows.map((r,i)=>i===editIdx?{...r,...form}:r);
      reCalcAndSave(newRows);
      setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 수정 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);
    } else {
      const saved=await dbInsert({...form,_src:"manual"});
      const newRow={...form,_id:saved?.id,_src:"manual"};
      const newRows=[...rows,newRow];
      reCalcAndSave(newRows);
      setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 등록 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);
    }
  }

  async function handleReject(rowIdx){
    const target=rows[rowIdx];
    if(target._id) await dbDelete(target._id);
    const newRows=rows.filter((_,i)=>i!==rowIdx);
    reCalcAndSave(newRows);
    setNotifs(prev=>prev.map(n=>(n.rowA===rowIdx||n.rowB===rowIdx)?{...n,resolved:true,unread:false}:n));
  }

  async function confirmCancel(){
    if(cancelTarget===null) return;
    const c=rows[cancelTarget];
    if(c._id) await dbDelete(c._id);
    const newRows=rows.filter((_,i)=>i!==cancelTarget);
    reCalcAndSave(newRows);
    setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`예약 취소 — ${c?.장소||""}`,desc:`${c?.날짜} ${c?.시작시간}~${c?.종료시간} 취소됨`,time:"방금",unread:true},...prev]);
    setCancelTarget(null);
  }

  async function cancelAllConflicts(){
    const cfIdxs=new Set(conflicts.flatMap(c=>[c.a,c.b]));
    const removed=rows.filter((_,i)=>cfIdxs.has(i));
    for(const r of removed){ if(r._id) await dbDelete(r._id); }
    const newRows=rows.filter((_,i)=>!cfIdxs.has(i));
    reCalcAndSave(newRows);
    setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`충돌 일괄 취소 — ${removed.length}건`,desc:`${[...new Set(removed.map(r=>r.장소))].join(", ")} 전체 제거됨`,time:"방금",unread:true},...prev]);
  }

  async function resetAllData(){
    if(!window.confirm("저장된 모든 예약 데이터를 초기화하시겠습니까?")) return;
    await supabase.from("bookings").delete().neq("id",0);
    reCalcAndSave([]);
    setNotifs([{id:0,type:"info",title:"초기화 완료",desc:"모든 데이터가 삭제되었습니다.",time:"방금",unread:false}]);
  }

  const monday=addDays(getThisWeekMonday(),weekOffset*7);
  const weekLabel=`${monday.getMonth()+1}월 ${fmtShort(monday)}(월) ~ ${fmtShort(addDays(monday,4))}(금)`;
  const unread=notifs.filter(n=>n.unread).length;
  const detectedStudios=[...new Set(rows.map(r=>r.장소).filter(Boolean))];
  const studioStats=detectedStudios.map(sid=>({id:sid,color:getColor(sid),count:rows.filter(r=>r.장소===sid).length,cf:conflicts.filter(c=>c.studio===sid).length}));
  const conflictIdxSet=new Set(conflicts.flatMap(c=>[c.a,c.b]));
  const filteredRows=filterStudio==="전체"?rows:rows.filter(r=>r.장소===filterStudio);

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",fontSize:14,background:"#f5f5f4",color:"#1c1c1a"}}>
      {bookingModal==="new"&&<BookingForm title="예약 등록" initial={null} onSave={f=>handleSave(f,null)} onClose={()=>setBookingModal(null)}/>}
      {typeof bookingModal==="number"&&<BookingForm title="예약 수정" initial={rows[bookingModal]} onSave={f=>handleSave(f,bookingModal)} onClose={()=>setBookingModal(null)}/>}
      {cancelTarget!==null&&<CancelModal row={rows[cancelTarget]} onClose={()=>setCancelTarget(null)} onConfirm={confirmCancel}/>}

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 20px",height:52,background:"#fff",borderBottom:"0.5px solid #e5e5e3",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8,fontWeight:600,fontSize:15}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:"#1D9E75"}}></div>
          스튜디오 방송편성 관리
          {saving&&<span style={{fontSize:11,color:"#aaa",fontWeight:400,marginLeft:4}}>저장 중...</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button style={btn} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
          <button style={{...btn,color:"#A32D2D",borderColor:"#F09595",fontSize:12}} onClick={resetAllData}>🗑 초기화</button>
          <button style={{...btn,position:"relative"}} onClick={()=>setTab("notifications")}>
            🔔{unread>0&&<span style={{position:"absolute",top:4,right:4,width:8,height:8,borderRadius:"50%",background:"#E24B4A",border:"1.5px solid #fff"}}></span>}
          </button>
        </div>
      </div>

      <div style={{display:"flex",flex:1,overflow:"hidden"}}>
        <nav style={{width:196,background:"#fff",borderRight:"0.5px solid #e5e5e3",padding:"14px 10px",overflowY:"auto",flexShrink:0}}>
          <div style={{fontSize:11,fontWeight:500,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 8px",marginBottom:6}}>메뉴</div>
          {[{id:"dashboard",icon:"📊",label:"대시보드"},{id:"upload",icon:"📤",label:"편성표 업로드"},{id:"schedule",icon:"📅",label:"스케줄 목록"},{id:"notifications",icon:"🔔",label:`알림${unread>0?` (${unread})`:""}`}].map(item=>(
            <div key={item.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 8px",borderRadius:8,fontSize:13,cursor:"pointer",color:tab===item.id?"#1c1c1a":"#666",fontWeight:tab===item.id?500:400,background:tab===item.id?"#f0f0ee":"transparent",marginBottom:2}} onClick={()=>setTab(item.id)}>
              <span>{item.icon}</span>{item.label}
            </div>
          ))}
          <div style={{margin:"12px 8px 6px",borderTop:"0.5px solid #e5e5e3"}}></div>
          <div style={{fontSize:11,fontWeight:500,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 8px",marginBottom:6}}>{studioStats.length>0?"스튜디오 현황":"스튜디오"}</div>
          {(studioStats.length>0?studioStats:STUDIOS.map(s=>({id:s.id,color:s.color,count:null,cf:0}))).map(st=>(
            <div key={st.id} style={{display:"flex",alignItems:"center",padding:"5px 8px",fontSize:12,color:st.count!=null?"#555":"#ccc",borderRadius:6,cursor:"pointer"}} onClick={()=>{setTab("schedule");setFilterStudio(st.id);}}>
              <div style={{width:7,height:7,borderRadius:"50%",background:st.color,marginRight:7,flexShrink:0}}></div>
              <div style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{st.id}</div>
              {st.count!=null&&<><div style={{fontSize:11,color:"#aaa",marginRight:2}}>{st.count}</div>{st.cf>0&&<span style={{...bd("red"),padding:"0 4px",fontSize:10}}>!</span>}</>}
            </div>
          ))}
          <div style={{margin:"12px 8px 10px",borderTop:"0.5px solid #e5e5e3"}}></div>
          <button style={{...btnP,width:"100%",justifyContent:"center",fontSize:12}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
        </nav>

        <main style={{flex:1,overflowY:"auto",padding:20,minWidth:0}}>
          {tab==="dashboard"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <div style={{fontSize:18,fontWeight:600}}>이번 주 스튜디오 스케줄</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <button style={{...btn,width:32,height:32,padding:0,justifyContent:"center"}} onClick={()=>setWeekOffset(w=>w-1)}>◀</button>
                  <span style={{fontSize:13,fontWeight:500,minWidth:180,textAlign:"center"}}>{weekLabel}</span>
                  <button style={{...btn,width:32,height:32,padding:0,justifyContent:"center"}} onClick={()=>setWeekOffset(w=>w+1)}>▶</button>
                  <button style={btn} onClick={()=>setTab("upload")}>📤 엑셀 업로드</button>
                  <button style={btnP} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:12}}>
                {[{label:"총 예약 건수",val:rows.length||"—",sub:rows.length?"전체 등록":"예약을 등록하세요"},{label:"운영 스튜디오",val:detectedStudios.length||"—",sub:detectedStudios.length?`${detectedStudios.length}개 운영 중`:"없음"},{label:"충돌 건수",val:conflicts.length||(rows.length?"0":"—"),sub:conflicts.length>0?"조율 필요":rows.length?"충돌 없음":"—",danger:conflicts.length>0},{label:"수동 예약",val:rows.filter(r=>r._src==="manual").length||(rows.length?"0":"—"),sub:"직접 등록"}].map((m,i)=>(
                  <div key={i} style={{background:"#f0f0ee",borderRadius:8,padding:"12px 14px"}}>
                    <div style={{fontSize:12,color:"#888",marginBottom:3}}>{m.label}</div>
                    <div style={{fontSize:20,fontWeight:500,color:m.danger?"#E24B4A":"#1c1c1a"}}>{m.val}</div>
                    <div style={{fontSize:11,color:m.danger?"#E24B4A":"#aaa",marginTop:1}}>{m.sub}</div>
                  </div>
                ))}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
                <span style={{fontSize:12,color:"#888"}}>💡 블록 클릭 시 상세보기·수정·취소 가능</span>
                <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  {STUDIOS.map(s=><span key={s.id} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#555"}}><span style={{width:10,height:10,borderRadius:3,background:s.color,display:"inline-block"}}></span>{s.id}</span>)}
                </span>
              </div>
              {rows.length===0?(
                <div style={{...panel,textAlign:"center",padding:60}}>
                  <div style={{fontSize:40,marginBottom:12}}>📊</div>
                  <div style={{fontSize:15,fontWeight:500,marginBottom:6}}>예약을 등록해주세요</div>
                  <div style={{fontSize:13,color:"#888",marginBottom:20,lineHeight:1.7}}>업로드 즉시 대시보드에 자동 반영 · Supabase DB에 영구 저장<br/>세로: 07:00~18:00 · 가로: 월~금</div>
                  <div style={{display:"flex",gap:10,justifyContent:"center"}}>
                    <button style={btn} onClick={()=>setTab("upload")}>📤 엑셀 일괄 업로드</button>
                    <button style={btnP} onClick={()=>setBookingModal("new")}>✚ 직접 예약 등록</button>
                  </div>
                </div>
              ):(
                <div style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,overflow:"hidden",marginBottom:12}}>
                  <WeeklyGrid rows={rows} conflicts={conflicts} monday={monday} onEdit={setBookingModal} onCancel={setCancelTarget}/>
                </div>
              )}
              {conflicts.length>0&&(
                <div style={{...panel,border:"0.5px solid #F09595",background:"#FFF8F8"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                    <div style={{fontSize:13,fontWeight:500,color:"#A32D2D"}}>⚠️ 중복 예약 충돌 {conflicts.length}건</div>
                    <button style={{...btnR,height:28,fontSize:12}} onClick={()=>{if(window.confirm(`충돌된 예약 ${[...new Set(conflicts.flatMap(c=>[c.a,c.b]))].length}건을 모두 취소하시겠습니까?`))cancelAllConflicts();}}>🗑 충돌 전체 취소</button>
                  </div>
                  {conflicts.map((cf,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"0.5px solid #F09595",fontSize:12}}>
                      <span style={{color:"#791F1F"}}>🏢 {cf.studio} · {cf.date}</span>
                      <span style={{color:"#888"}}>{cf.timeA} ↔ {cf.timeB}</span>
                      <div style={{display:"flex",gap:6}}>
                        <button style={{...btn,height:26,fontSize:11}} onClick={()=>setBookingModal(cf.a)}>A 수정</button>
                        <button style={{...btn,height:26,fontSize:11}} onClick={()=>setBookingModal(cf.b)}>B 수정</button>
                        <button style={{...btnR,height:26,fontSize:11}} onClick={()=>setCancelTarget(cf.a)}>A 취소</button>
                        <button style={{...btnR,height:26,fontSize:11}} onClick={()=>setCancelTarget(cf.b)}>B 취소</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}

          {tab==="upload"&&(
            <>
              <div style={{fontSize:18,fontWeight:600,marginBottom:16}}>편성표 업로드</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
                <div style={panel}>
                  <div style={{fontSize:14,fontWeight:500,marginBottom:4}}>📊 엑셀 일괄 업로드</div>
                  <div style={{fontSize:12,color:"#888",marginBottom:14}}>VBA로 생성한 방송편성표 엑셀 → Supabase DB 저장</div>
                  <div style={{border:uploadState==="done"?"1.5px solid #1D9E75":"1px dashed #ccc",borderRadius:10,padding:24,textAlign:"center",cursor:"pointer",marginBottom:10}} onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();handleFile({target:{files:e.dataTransfer.files}});}}>
                    <div style={{fontSize:28,marginBottom:6}}>{uploadState==="done"?"✅":uploadState==="parsing"?"⏳":uploadState==="error"?"❌":"📂"}</div>
                    <div style={{fontSize:13,fontWeight:500,marginBottom:2}}>{uploadState==="idle"?"파일 클릭 또는 드래그":uploadState==="parsing"?"처리 중...":uploadState==="done"?fileName:"업로드 실패"}</div>
                    <div style={{fontSize:11,color:"#aaa"}}>B열(장소) 자동 인식 · .xlsx .xls</div>
                    <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleFile}/>
                  </div>
                  {parseLog.map((l,i)=><div key={i} style={{fontSize:12,lineHeight:1.9,color:l.startsWith("⚠")||l.startsWith("❌")?"#A32D2D":"#0F6E56"}}>{l}</div>)}
                </div>
                <div style={panel}>
                  <div style={{fontSize:14,fontWeight:500,marginBottom:4}}>✚ 스튜디오 직접 예약</div>
                  <div style={{fontSize:12,color:"#888",marginBottom:14}}>스튜디오·날짜·시간을 직접 입력해 Supabase DB에 저장</div>
                  <div style={{display:"flex",flexDirection:"column",gap:7}}>
                    {STUDIOS.map(s=>(
                      <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",border:"0.5px solid #e5e5e3",borderRadius:8,fontSize:12}}>
                        <span style={{display:"flex",alignItems:"center",gap:7}}><span style={{width:8,height:8,borderRadius:"50%",background:s.color}}></span>{s.id}</span>
                        <button style={{...btn,height:26,fontSize:11}} onClick={()=>setBookingModal("new")}>예약하기</button>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </>
          )}

          {tab==="schedule"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontSize:18,fontWeight:600}}>스케줄 목록</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <select style={{height:34,padding:"0 10px",borderRadius:8,border:"0.5px solid #ccc",fontSize:12,background:"#fff"}} value={filterStudio} onChange={e=>setFilterStudio(e.target.value)}>
                    <option>전체</option>
                    {STUDIOS.map(s=><option key={s.id}>{s.id}</option>)}
                  </select>
                  <span style={{fontSize:13,color:"#888"}}>{filteredRows.length}건</span>
                  <button style={btnP} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>
              {filteredRows.length===0?(
                <div style={{...panel,textAlign:"center",padding:40}}>
                  <div style={{fontSize:13,color:"#888",marginBottom:12}}>등록된 예약이 없습니다.</div>
                  <div style={{display:"flex",gap:8,justifyContent:"center"}}>
                    <button style={btn} onClick={()=>setTab("upload")}>📤 엑셀 업로드</button>
                    <button style={btnP} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                  </div>
                </div>
              ):(
                <div style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,overflow:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",minWidth:920}}>
                    <thead>
                      <tr style={{background:"#f8f8f6",borderBottom:"0.5px solid #e5e5e3"}}>
                        {["날짜","요일","구분","장소","내용","강사명","시작","종료","길이","출처","상태","수정","취소"].map(h=>(
                          <th key={h} style={{padding:"8px 10px",fontSize:11,fontWeight:500,color:"#888",textAlign:"left",borderRight:"0.5px solid #e5e5e3",whiteSpace:"nowrap"}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row,i)=>{
                        const realIdx=rows.indexOf(row);
                        const isCf=conflictIdxSet.has(realIdx);
                        const noSt=!row.장소;
                        const td=(children,extra={})=><td style={{padding:"6px 10px",fontSize:11,borderRight:"0.5px solid #f0f0ee",...extra}}>{children}</td>;
                        return(
                          <tr key={i} style={{borderBottom:"0.5px solid #f0f0ee",background:isCf?"#FFF8F8":noSt?"#FFFBF0":"#fff"}}>
                            {td(row.날짜,{whiteSpace:"nowrap"})}
                            {td(row.요일)}
                            {td(row.구분)}
                            {td(row.장소?<span style={{display:"inline-flex",alignItems:"center",gap:5}}><span style={{width:6,height:6,borderRadius:"50%",background:getColor(row.장소),flexShrink:0}}></span>{row.장소}</span>:<span style={bd("amber")}>미입력</span>)}
                            {td(row.내용||row.주제,{maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"})}
                            {td(row.강사명,{whiteSpace:"nowrap"})}
                            {td(row.시작시간)}
                            {td(row.종료시간)}
                            {td(row.길이)}
                            {td(<span style={bd(row._src==="manual"?"blue":"gray")}>{row._src==="manual"?"직접":"엑셀"}</span>)}
                            {td(isCf?<span style={bd("red")}>충돌</span>:noSt?<span style={bd("amber")}>장소미정</span>:<span style={bd("green")}>확정</span>)}
                            <td style={{padding:"6px 10px",fontSize:11,borderRight:"0.5px solid #f0f0ee"}}><button style={{...btn,height:26,padding:"0 10px",fontSize:11,color:"#185FA5",borderColor:"#B5D4F4"}} onClick={()=>setBookingModal(realIdx)}>✏ 수정</button></td>
                            <td style={{padding:"6px 10px",fontSize:11}}><button style={{...btnR,height:26,padding:"0 10px",fontSize:11}} onClick={()=>setCancelTarget(realIdx)}>✕ 취소</button></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}

          {tab==="notifications"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontSize:18,fontWeight:600}}>알림</div>
                <button style={btn} onClick={()=>setNotifs(notifs.map(n=>({...n,unread:false})))}>모두 읽음</button>
              </div>
              {notifs.map(n=>(
                <div key={n.id} style={{background:n.unread?"#FAFFF9":"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,padding:"12px 14px",display:"flex",gap:10,marginBottom:8,opacity:n.resolved?0.55:1}}>
                  <div style={{width:32,height:32,borderRadius:8,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,background:n.type==="conflict"?"#FCEBEB":n.type==="cancel"?"#FFF0F0":n.type==="manual"?"#E6F1FB":"#E1F5EE",color:n.type==="conflict"?"#A32D2D":n.type==="cancel"?"#A32D2D":n.type==="manual"?"#185FA5":"#0F6E56"}}>
                    {n.type==="conflict"?"⚠️":n.type==="cancel"?"🚫":n.type==="manual"?"✏️":"📤"}
                  </div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:500,marginBottom:3}}>{n.title}</div>
                    <div style={{fontSize:12,color:"#888",lineHeight:1.6}}>{n.desc}</div>
                    {n.resolved&&<div style={{fontSize:11,color:"#0F6E56",marginTop:4}}>✓ 처리 완료</div>}
                    {n.type==="conflict"&&!n.resolved&&(
                      <div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>
                        <button style={{...btn,height:28,fontSize:11}} onClick={()=>setBookingModal(n.rowA)}>A 수정</button>
                        <button style={{...btn,height:28,fontSize:11}} onClick={()=>setBookingModal(n.rowB)}>B 수정</button>
                        <button style={{...btnR,height:28,fontSize:11}} onClick={()=>handleReject(n.rowA)}>A 반려</button>
                        <button style={{...btnR,height:28,fontSize:11}} onClick={()=>handleReject(n.rowB)}>B 반려</button>
                      </div>
                    )}
                    {n.type==="upload"&&(
                      <div style={{marginTop:8,display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{...bd("green"),fontSize:11}}>✓ DB 저장 완료</span>
                        <button style={{...btn,height:26,fontSize:11}} onClick={()=>setNotifs(prev=>prev.map(x=>x.id===n.id?{...x,unread:false}:x))}>확인</button>
                      </div>
                    )}
                  </div>
                  <div style={{fontSize:11,color:"#aaa",flexShrink:0,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
                    {n.unread&&!n.resolved&&<span style={{...bd("red"),padding:"1px 6px",fontSize:10}}>new</span>}
                    {n.time}
                  </div>
                </div>
              ))}
            </>
          )}
        </main>
      </div>
    </div>
  );
}