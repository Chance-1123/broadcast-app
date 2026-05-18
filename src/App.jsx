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
  { id:"18층 대강당",    color:"#C2410C" },
  { id:"대전",           color:"#0891B2" },
  { id:"광주",           color:"#0C447C" },
  { id:"구미",           color:"#993C1D" },
];
const ALIASES = {
  "지하":"스튜디오(지하)","스튜디오지하":"스튜디오(지하)","스튜디오(지하)":"스튜디오(지하)",
  "17층1호":"17층1호","17층 1호":"17층1호","17-1":"17층1호",
  "17층2호":"17층2호","17층 2호":"17층2호","17-2":"17층2호",
  "17층3호":"17층3호","17층 3호":"17층3호","17-3":"17층3호",
  "온택트룸2":"온택트룸2","온택트룸 2":"온택트룸2","온택2":"온택트룸2",
  "온택트룸4":"온택트룸4","온택트룸 4":"온택트룸4","온택4":"온택트룸4",
  "온택트룸6":"온택트룸6","온택트룸 6":"온택트룸6","온택6":"온택트룸6",
  "18층 대강당":"18층 대강당","18층대강당":"18층 대강당","대강당":"18층 대강당",
  "대전":"대전","광주":"광주","구미":"구미",
};
const DAYS=["월","화","수","목","금"];
const HOURS=Array.from({length:12},(_,i)=>`${String(i+7).padStart(2,"0")}:00`);

function excelDateToStr(val){if(!val&&val!==0)return "";const s=String(val).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;if(typeof val==="number"&&val>1000){const d=new Date(Math.round((val-25569)*86400*1000));return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;}return s;}
function excelTimeToStr(val){if(!val&&val!==0)return "";const s=String(val).trim();if(/^\d{1,2}:\d{2}$/.test(s))return s.padStart(5,"0");if(typeof val==="number"&&val>=0&&val<1){const m=Math.round(val*24*60);return `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;}return s;}
function calcLen(s,e){const sm=toMin(s),em=toMin(e);if(sm===null||em===null||em<=sm)return "";const d=em-sm;return `${String(Math.floor(d/60)).padStart(2,"0")}:${String(d%60).padStart(2,"0")}`;}
function normStudio(raw){if(!raw)return "";const t=String(raw).trim();return ALIASES[t]||ALIASES[t.replace(/\s/g,"")]||t;}
function toMin(t){if(!t)return null;const p=String(t).split(":");if(p.length<2)return null;const h=parseInt(p[0]),m=parseInt(p[1]);return isNaN(h)||isNaN(m)?null:h*60+m;}
function getThisWeekMonday(){const d=new Date();const dow=d.getDay()||7;d.setDate(d.getDate()-dow+1);return d;}
function addDays(d,n){const r=new Date(d);r.setDate(r.getDate()+n);return r;}
function fmtShort(d){return `${d.getMonth()+1}/${d.getDate()}`;}
function fmtFull(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
function getColor(id){return STUDIOS.find(s=>s.id===id)?.color||"#888";}
function detectConflicts(rows){const res=[];for(let i=0;i<rows.length;i++)for(let j=i+1;j<rows.length;j++){const a=rows[i],b=rows[j];if(!a.장소||a.장소!==b.장소||a.날짜!==b.날짜)continue;const aS=toMin(a.시작시간),aE=toMin(a.종료시간),bS=toMin(b.시작시간),bE=toMin(b.종료시간);if(aS!==null&&aE!==null&&bS!==null&&bE!==null&&aS<bE&&bS<aE)res.push({a:i,b:j,studio:a.장소,date:a.날짜,timeA:`${a.시작시간}~${a.종료시간}`,timeB:`${b.시작시간}~${b.종료시간}`});}return res;}
function downloadSchedule(rows){if(rows.length===0){alert("다운로드할 데이터가 없습니다.");return;}const data=rows.map(r=>({구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이}));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"방송편성표");ws['!cols']=[{wch:8},{wch:14},{wch:24},{wch:30},{wch:14},{wch:12},{wch:6},{wch:10},{wch:10},{wch:8}];XLSX.writeFile(wb,`방송편성표_${new Date().toISOString().slice(0,10)}.xlsx`);}

const bd=(c)=>{const m={red:{bg:"#FCEBEB",fg:"#A32D2D"},green:{bg:"#E1F5EE",fg:"#0F6E56"},blue:{bg:"#E6F1FB",fg:"#185FA5"},amber:{bg:"#FAEEDA",fg:"#854F0B"},gray:{bg:"#f0f0ee",fg:"#666"}};const x=m[c]||m.gray;return{display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:16,fontWeight:500,background:x.bg,color:x.fg};};
const btn={display:"inline-flex",alignItems:"center",gap:5,padding:"0 14px",height:34,borderRadius:8,border:"0.5px solid #ccc",fontSize:20,cursor:"pointer",background:"#fff",color:"#333",fontFamily:"inherit"};
const btnP={...btn,background:"#1D9E75",borderColor:"#1D9E75",color:"#fff"};
const btnR={...btn,background:"#FCEBEB",borderColor:"#F09595",color:"#A32D2D"};
const btnB={...btn,background:"#E6F1FB",borderColor:"#B5D4F4",color:"#185FA5"};
const panel={background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,padding:"16px 18px",marginBottom:12};
const inp={width:"100%",height:36,padding:"0 10px",borderRadius:8,border:"0.5px solid #ccc",fontSize:20,fontFamily:"inherit",background:"#fff",color:"#1c1c1a",boxSizing:"border-box"};
const lbl={fontSize:18,fontWeight:500,color:"#555",marginBottom:4,display:"block"};

function BookingForm({initial,onSave,onClose,title}){
  const [form,setForm]=useState(initial||{장소:"",날짜:"",구분:"1학기",주제:"",내용:"",강사명:"",시작시간:"09:00",종료시간:"18:00",요일:""});
  const [customStudio,setCustomStudio]=useState(initial?.장소&&!STUDIOS.find(s=>s.id===initial.장소)?initial.장소:"");
  const [useCustom,setUseCustom]=useState(!!(initial?.장소&&!STUDIOS.find(s=>s.id===initial.장소)));
  const [err,setErr]=useState("");
  function set(k,v){let u={...form,[k]:v};if(k==="날짜"&&v){const d=new Date(v);u.요일=["일","월","화","수","목","금","토"][d.getDay()];}setForm(u);}
  const len=calcLen(form.시작시간,form.종료시간);
  function save(){
    const 장소=useCustom?customStudio.trim():form.장소;
    if(!장소)return setErr("스튜디오를 선택하거나 직접 입력해주세요");
    if(!form.날짜)return setErr("날짜를 입력해주세요");
    if(toMin(form.시작시간)>=toMin(form.종료시간))return setErr("종료 시간이 시작 시간보다 늦어야 합니다");
    onSave({...form,장소,길이:len});onClose();
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"#fff",borderRadius:16,width:500,maxHeight:"92vh",overflowY:"auto",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 20px",borderBottom:"0.5px solid #e5e5e3",position:"sticky",top:0,background:"#fff"}}>
          <div style={{fontSize:24,fontWeight:600}}>{title}</div>
          <button style={{...btn,height:28,padding:"0 10px"}} onClick={onClose}>✕</button>
        </div>
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",gap:12}}>
          <div>
            <span style={lbl}>스튜디오 *</span>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6,marginBottom:8}}>
              {STUDIOS.map(s=>{
                const active=!useCustom&&form.장소===s.id;
                return(
                  <label key={s.id} style={{display:"flex",alignItems:"center",gap:7,padding:"7px 10px",border:`1.5px solid ${active?s.color:"#e5e5e3"}`,borderRadius:8,cursor:"pointer",background:active?`${s.color}10`:"#fff",fontSize:18}}>
                    <input type="radio" name="bst" style={{display:"none"}} checked={active} onChange={()=>{set("장소",s.id);setUseCustom(false);}}/>
                    <span style={{width:9,height:9,borderRadius:"50%",background:s.color,flexShrink:0}}></span>
                    {s.id}{active&&<span style={{marginLeft:"auto",color:s.color,fontWeight:700}}>✓</span>}
                  </label>
                );
              })}
            </div>
            {/* 직접 입력 */}
            <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",border:`1.5px solid ${useCustom?"#378ADD":"#e5e5e3"}`,borderRadius:8,background:useCustom?"#F0F6FF":"#f8f8f6",cursor:"pointer"}} onClick={()=>{setUseCustom(true);set("장소","");}}>
              <input type="radio" name="bst" style={{display:"none"}} checked={useCustom} onChange={()=>{setUseCustom(true);set("장소","");}}/>
              <span style={{fontSize:14,fontWeight:600,color:useCustom?"#378ADD":"#888",flexShrink:0,whiteSpace:"nowrap"}}>✏ 직접 입력</span>
              {useCustom?(
                <input type="text" style={{...inp,flex:1,border:"none",background:"transparent",padding:"0 6px",fontSize:16}} placeholder="강의장명 직접 입력 (예: A동 101호)" value={customStudio} onChange={e=>setCustomStudio(e.target.value)} onClick={e=>e.stopPropagation()} autoFocus/>
              ):(
                <span style={{fontSize:14,color:"#aaa",marginLeft:4}}>목록에 없는 강의장 직접 입력</span>
              )}
            </div>
          </div>
          <div><span style={lbl}>구분 *</span>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {["방송준비","1학기","2학기","취업","기타"].map(g=>(
                <label key={g} style={{display:"flex",alignItems:"center",gap:5,padding:"6px 12px",border:`1.5px solid ${form.구분===g?"#378ADD":"#e5e5e3"}`,borderRadius:8,cursor:"pointer",fontSize:18,background:form.구분===g?"#F0F6FF":"#fff"}}>
                  <input type="radio" name="bgb" style={{display:"none"}} checked={form.구분===g} onChange={()=>set("구분",g)}/>{g}
                </label>
              ))}
            </div>
          </div>
          <div><span style={lbl}>날짜 *</span><input type="date" style={inp} value={form.날짜} onChange={e=>set("날짜",e.target.value)}/>{form.요일&&<div style={{fontSize:16,color:"#1D9E75",marginTop:3}}>{form.요일}요일</div>}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><span style={lbl}>시작 시간 *</span><input type="time" style={inp} value={form.시작시간} onChange={e=>set("시작시간",e.target.value)}/></div>
            <div><span style={lbl}>종료 시간 *</span><input type="time" style={inp} value={form.종료시간} onChange={e=>set("종료시간",e.target.value)}/></div>
          </div>
          {len&&<div style={{fontSize:18,color:"#1D9E75",fontWeight:500}}>⏱ 총 {len}</div>}
          <div><span style={lbl}>주제</span><input type="text" style={inp} placeholder="예: 코딩 Live Django트랙" value={form.주제} onChange={e=>set("주제",e.target.value)}/></div>
          <div><span style={lbl}>내용</span><input type="text" style={inp} placeholder="예: Django : DRF 1" value={form.내용} onChange={e=>set("내용",e.target.value)}/></div>
          <div><span style={lbl}>강사명</span><input type="text" style={inp} placeholder="예: 김준호 강사" value={form.강사명} onChange={e=>set("강사명",e.target.value)}/></div>
          {err&&<div style={{fontSize:18,color:"#E24B4A",background:"#FFF0F0",padding:"8px 12px",borderRadius:8}}>{err}</div>}
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
      <div style={{background:"#fff",borderRadius:16,width:380,padding:"24px",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{fontSize:24,fontWeight:600,marginBottom:12}}>예약을 취소하시겠습니까?</div>
        <div style={{fontSize:20,color:"#555",marginBottom:12,lineHeight:1.9,background:"#f8f8f6",borderRadius:8,padding:"10px 14px"}}>
          <div><b>스튜디오:</b> {row?.장소}</div><div><b>날짜:</b> {row?.날짜} {row?.요일&&`(${row.요일}요일)`}</div>
          <div><b>시간:</b> {row?.시작시간} ~ {row?.종료시간}</div><div><b>내용:</b> {row?.내용||row?.주제||"-"}</div>
        </div>
        <div style={{background:"#FFF8F8",border:"0.5px solid #F09595",borderRadius:8,padding:"8px 12px",fontSize:18,color:"#A32D2D",marginBottom:16}}>⚠ 취소된 예약은 복구할 수 없습니다.</div>
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
    <div style={{position:"fixed",top:Math.min(pos.y+10,window.innerHeight-340),left:Math.min(pos.x+10,window.innerWidth-280),zIndex:3000,background:"#fff",borderRadius:12,border:"0.5px solid #e5e5e3",boxShadow:"0 8px 28px rgba(0,0,0,0.18)",width:260,padding:"14px 16px"}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{display:"flex",alignItems:"center",gap:6,fontWeight:600,fontSize:20,color}}><span style={{width:9,height:9,borderRadius:"50%",background:color}}></span>{row.장소}</span>
        <button style={{background:"none",border:"none",cursor:"pointer",fontSize:24,color:"#aaa"}} onClick={onClose}>✕</button>
      </div>
      <div style={{fontSize:18,lineHeight:2,color:"#444",borderBottom:"0.5px solid #f0f0ee",paddingBottom:10,marginBottom:10}}>
        {row._conflict&&<div style={{color:"#E24B4A",fontWeight:600,marginBottom:4}}>⚠ 충돌</div>}
        <div><b>구분:</b> {row.구분}</div><div><b>내용:</b> {row.내용||row.주제||"-"}</div>
        {row.강사명&&<div><b>강사:</b> {row.강사명}</div>}
        <div><b>날짜:</b> {row.날짜} {row.요일&&`(${row.요일})`}</div><div><b>시간:</b> {row.시작시간} ~ {row.종료시간}</div>
      </div>
      <div style={{display:"flex",gap:6}}>
        <button style={{...btn,flex:1,height:30,fontSize:18,justifyContent:"center"}} onClick={()=>{onEdit(idx);onClose();}}>✏ 수정</button>
        <button style={{...btnR,flex:1,height:30,fontSize:18,justifyContent:"center"}} onClick={()=>{onCancel(idx);onClose();}}>✕ 취소</button>
      </div>
    </div>
  );
}

function WeeklyGrid({rows,activeStudios,conflicts,monday,onEdit,onCancel}){
  const [popup,setPopup]=useState(null);
  const dayDates=DAYS.map((_,i)=>fmtFull(addDays(monday,i)));
  const conflictIdxs=new Set();
  conflicts.forEach(c=>{conflictIdxs.add(c.a);conflictIdxs.add(c.b);});

  // 스튜디오 필터 적용 (activeStudios가 비어있으면 전체)
  const displayRows=(!activeStudios||activeStudios.size===0)
    ?rows
    :rows.filter(r=>activeStudios.has(r.장소));

  function getBlocksInHour(dayDate,hourStr){
    const hMin=toMin(hourStr);
    return displayRows.map((r,i)=>({...r,idx:rows.indexOf(r)})).filter(r=>{
      if(r.날짜!==dayDate||!r.시작시간||!r.장소)return false;
      const s=toMin(r.시작시간),e=toMin(r.종료시간);
      if(s===null||e===null)return false;
      return s<hMin+60&&e>hMin;
    }).sort((a,b)=>toMin(a.시작시간)-toMin(b.시작시간));
  }
  return(
    <div style={{position:"relative",width:"100%"}} onClick={()=>setPopup(null)}>
      {popup&&<BlockPopup row={popup.row} idx={popup.idx} pos={{x:popup.x,y:popup.y}} onClose={()=>setPopup(null)} onEdit={onEdit} onCancel={onCancel}/>}
      <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
        <thead>
          <tr style={{position:"sticky",top:0,zIndex:2}}>
            <th style={{width:56,padding:"8px 4px",background:"#f8f8f6",borderBottom:"2px solid #e5e5e3",borderRight:"0.5px solid #e5e5e3",fontSize:16,fontWeight:600,color:"#888",textAlign:"center"}}>시간</th>
            {DAYS.map((day,di)=>{
              const isToday=fmtFull(new Date())===dayDates[di];
              const cnt=displayRows.filter(r=>r.날짜===dayDates[di]&&r.장소).length;
              return(
                <th key={day} style={{padding:"8px 4px",background:isToday?"#F0FBF6":"#f8f8f6",borderBottom:"2px solid #e5e5e3",borderRight:"0.5px solid #e5e5e3",fontSize:18,fontWeight:700,color:isToday?"#1D9E75":"#222",textAlign:"center"}}>
                  <div>{day}요일</div>
                  <div style={{fontSize:15,fontWeight:400,color:isToday?"#1D9E75":"#aaa"}}>{fmtShort(addDays(monday,di))}</div>
                  {cnt>0&&<div style={{fontSize:14,color:isToday?"#1D9E75":"#aaa"}}>{cnt}건</div>}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {HOURS.map((hour,hi)=>(
            <tr key={hour} style={{background:hi%2===0?"#fff":"#fafafa"}}>
              <td style={{padding:"5px 4px",borderBottom:"0.5px solid #f0f0ee",borderRight:"0.5px solid #e5e5e3",textAlign:"center",verticalAlign:"top",background:hi%2===0?"#fafaf9":"#f4f4f3"}}>
                <span style={{fontSize:16,fontWeight:600,color:"#aaa"}}>{hour}</span>
              </td>
              {DAYS.map((day,di)=>{
                const isToday=fmtFull(new Date())===dayDates[di];
                const allBlocks=getBlocksInHour(dayDates[di],hour);
                const hMin=toMin(hour);
                const startBlocks=allBlocks.filter(b=>toMin(b.시작시간)>=hMin&&toMin(b.시작시간)<hMin+60);
                const continueBlocks=allBlocks.filter(b=>toMin(b.시작시간)<hMin);
                return(
                  <td key={day} style={{padding:"3px 4px",borderBottom:"0.5px solid #f0f0ee",borderRight:"0.5px solid #e5e5e3",verticalAlign:"top",background:isToday?"rgba(29,158,117,0.02)":"transparent"}}>
                    {allBlocks.length===0&&<div style={{height:28,display:"flex",alignItems:"center",justifyContent:"center",color:"#f0f0ee",fontSize:15}}>·</div>}
                    {continueBlocks.map((b,bi)=>{
                      const isCf=conflictIdxs.has(b.idx);
                      const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                      const color=isCf?"#E24B4A":isPrep?"#999":getColor(b.장소);
                      const bg=isPrep?"rgba(180,180,180,0.1)":isCf?"rgba(226,74,74,0.08)":`${getColor(b.장소)}08`;
                      const content=b.내용||b.주제||"";
                      const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                      const dur=calcLen(b.시작시간,b.종료시간);
                      return(
                        <div key={`c${bi}`}
                          style={{background:bg,border:`1.5px solid ${color}80`,borderLeft:`3px solid ${color}`,borderRadius:5,padding:"4px 6px",marginBottom:3,cursor:"pointer",opacity:isPrep?0.7:0.85,transition:"box-shadow 0.15s"}}
                          onClick={e=>{e.stopPropagation();setPopup({row:{...b,_conflict:isCf},idx:b.idx,x:e.clientX,y:e.clientY});}}
                          onMouseEnter={e=>e.currentTarget.style.boxShadow=`0 2px 8px ${color}40`}
                          onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}
                        >
                          <div style={{fontSize:11,color,fontWeight:600,marginBottom:1}}>↑ 계속</div>
                          <div style={{fontSize:14,fontWeight:700,color,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{b.장소}</div>
                          <div style={{fontSize:14,fontWeight:isPrep?400:600,color:isPrep?"#aaa":"#333",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{isPrep?"📋 방송준비":shortContent||b.주제||b.내용}</div>
                          {!isPrep&&b.강사명&&<div style={{fontSize:12,color:"#888",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>👤 {b.강사명}</div>}
                          <div style={{fontSize:12,color:"#aaa"}}>{b.시작시간}~{b.종료시간}{dur?` (${dur})`:""}</div>
                        </div>
                      );
                    })}
                    {startBlocks.map((b,bi)=>{
                      const isCf=conflictIdxs.has(b.idx);
                      const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                      const color=isCf?"#E24B4A":isPrep?"#999":getColor(b.장소);
                      const bg=isPrep?"rgba(180,180,180,0.1)":isCf?"rgba(226,74,74,0.08)":`${getColor(b.장소)}12`;
                      const content=b.내용||b.주제||"";
                      const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                      const dur=calcLen(b.시작시간,b.종료시간);
                      return(
                        <div key={bi} style={{background:bg,border:`1.5px solid ${color}`,borderLeft:`3px solid ${color}`,borderRadius:5,padding:"4px 5px",marginBottom:3,cursor:"pointer",transition:"box-shadow 0.15s"}}
                          onClick={e=>{e.stopPropagation();setPopup({row:{...b,_conflict:isCf},idx:b.idx,x:e.clientX,y:e.clientY});}}
                          onMouseEnter={e=>e.currentTarget.style.boxShadow=`0 2px 8px ${color}40`}
                          onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                          {isCf&&<div style={{fontSize:12,color:"#E24B4A",fontWeight:700}}>⚠ 충돌</div>}
                          <div style={{fontSize:14,fontWeight:700,color,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{b.장소}</div>
                          <div style={{fontSize:14,fontWeight:isPrep?400:600,color:isPrep?"#aaa":"#333",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{isPrep?"📋 방송준비":shortContent||b.주제||b.내용}</div>
                          {!isPrep&&b.강사명&&<div style={{fontSize:12,color:"#888",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>👤 {b.강사명}</div>}
                          <div style={{fontSize:12,color:"#aaa"}}>{b.시작시간}~{b.종료시간}{dur?` (${dur})`:""}</div>
                          {!isPrep&&<div style={{display:"flex",gap:2,marginTop:3}}>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:`0.5px solid ${color}`,borderRadius:3,fontSize:12,cursor:"pointer",padding:"1px 0",color,fontWeight:700}} onClick={e=>{e.stopPropagation();onEdit(b.idx);}}>✏ 수정</button>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:"0.5px solid #F09595",borderRadius:3,fontSize:12,cursor:"pointer",padding:"1px 0",color:"#A32D2D",fontWeight:700}} onClick={e=>{e.stopPropagation();onCancel(b.idx);}}>✕ 취소</button>
                          </div>}
                        </div>
                      );
                    })}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function LiveBanner({rows}){
  const [rollIdx,setRollIdx]=useState(0);
  const nowMin=new Date().getHours()*60+new Date().getMinutes();
  const live=rows.filter(r=>{const s=toMin(r.시작시간),e=toMin(r.종료시간);return s!==null&&e!==null&&s<=nowMin&&e>nowMin&&r.장소;});
  useEffect(()=>{if(live.length<=1)return;const t=setInterval(()=>setRollIdx(i=>(i+1)%live.length),2800);return()=>clearInterval(t);},[live.length]);
  if(live.length===0)return(
    <div style={{display:"flex",alignItems:"center",gap:10,background:"#f8f8f6",border:"0.5px solid #e5e5e3",borderRadius:10,padding:"10px 16px",marginBottom:10}}>
      <div style={{width:8,height:8,borderRadius:"50%",background:"#ccc"}}></div>
      <span style={{fontSize:18,color:"#aaa"}}>현재 진행 중인 라이브 방송이 없습니다</span>
      <span style={{marginLeft:"auto",fontSize:16,color:"#ccc"}}>{String(new Date().getHours()).padStart(2,"0")}:{String(new Date().getMinutes()).padStart(2,"0")} 기준</span>
    </div>
  );
  const cur=live[rollIdx%live.length];
  const color=getColor(cur.장소);
  const content=cur.내용||cur.주제||"";
  const shortContent=content.includes(":")?content.split(":").pop().trim():content;
  return(
    <div style={{display:"flex",alignItems:"center",gap:12,background:`${color}10`,border:`1.5px solid ${color}40`,borderRadius:10,padding:"10px 16px",marginBottom:10,position:"relative"}}>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}`}</style>
      <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
        <div style={{width:8,height:8,borderRadius:"50%",background:"#E24B4A",animation:"pulse 1.2s infinite"}}></div>
        <span style={{fontSize:16,fontWeight:700,color:"#E24B4A",letterSpacing:"0.5px"}}>LIVE</span>
      </div>
      <div style={{width:1,height:28,background:`${color}40`,flexShrink:0}}></div>
      <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
        <span style={{width:8,height:8,borderRadius:"50%",background:color}}></span>
        <span style={{fontSize:18,fontWeight:700,color}}>{cur.장소}</span>
      </div>
      <div style={{flex:1,minWidth:0}}>
        <span style={{fontSize:18,fontWeight:600,color:"#222",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",display:"block"}}>{shortContent||content}</span>
        {cur.강사명&&<span style={{fontSize:16,color:"#666"}}>👤 {cur.강사명}</span>}
      </div>
      <div style={{flexShrink:0,textAlign:"right"}}>
        <div style={{fontSize:16,fontWeight:600,color}}>{cur.시작시간}~{cur.종료시간}</div>
        {cur.길이&&<div style={{fontSize:15,color:"#aaa"}}>{cur.길이}</div>}
      </div>
      {live.length>1&&<div style={{display:"flex",gap:3,flexShrink:0}}>{live.map((_,i)=><div key={i} style={{width:5,height:5,borderRadius:"50%",background:i===rollIdx%live.length?color:"#ddd",transition:"background 0.3s"}}></div>)}</div>}
      {live.length>1&&<div style={{position:"absolute",right:14,top:6,fontSize:15,color:"#aaa"}}>{rollIdx%live.length+1}/{live.length}</div>}
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
  const [filterGubun,setFilterGubun]=useState("전체");
  const [filterDate,setFilterDate]=useState("");
  const [activeStudios,setActiveStudios]=useState(new Set()); // 빈 Set = 전체 표시
  const [fileName,setFileName]=useState("");
  const [saving,setSaving]=useState(false);
  const fileRef=useRef();

  function toggleStudio(id){
    setActiveStudios(prev=>{
      const next=new Set(prev);
      if(next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  useEffect(()=>{
    async function load(){
      const {data,error}=await supabase.from("bookings").select("*").order("created_at",{ascending:true});
      if(!error&&data){const mapped=data.map(r=>({_id:r.id,구분:r.구분||"",장소:r.장소||"",주제:r.주제||"",내용:r.내용||"",강사명:r.강사명||"",날짜:r.날짜||"",요일:r.요일||"",시작시간:r.시작시간||"",종료시간:r.종료시간||"",길이:r.길이||"",_src:r.src||"manual"}));setRows(mapped);setConflicts(detectConflicts(mapped));}
    }
    load();
  },[]);

  function reCalc(newRows){setRows(newRows);setConflicts(detectConflicts(newRows));}
  async function dbInsert(row){setSaving(true);const {data}=await supabase.from("bookings").insert([{구분:row.구분,장소:row.장소,주제:row.주제,내용:row.내용,강사명:row.강사명,날짜:row.날짜,요일:row.요일,시작시간:row.시작시간,종료시간:row.종료시간,길이:row.길이,src:row._src||"manual"}]).select();setSaving(false);return data?.[0];}
  async function dbUpdate(id,row){setSaving(true);await supabase.from("bookings").update({구분:row.구분,장소:row.장소,주제:row.주제,내용:row.내용,강사명:row.강사명,날짜:row.날짜,요일:row.요일,시작시간:row.시작시간,종료시간:row.종료시간,길이:row.길이}).eq("id",id);setSaving(false);}
  async function dbDelete(id){await supabase.from("bookings").delete().eq("id",id);}

  function handleFile(e){
    const file=e.target.files[0];if(!file)return;
    setFileName(file.name);setUploadState("parsing");
    const reader=new FileReader();
    reader.onload=async(evt)=>{
      try{
        const wb=XLSX.read(evt.target.result,{type:"binary",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
        const parsed=raw.map(r=>{const 날짜=excelDateToStr(r["날짜"]||r["F"]||"");const 시작=excelTimeToStr(r["시작시간"]||r["H"]||"");const 종료=excelTimeToStr(r["종료시간"]||r["I"]||"");return{구분:String(r["구분"]||r["A"]||""),장소:normStudio(r["장소"]||r["B"]||""),주제:String(r["주제"]||r["C"]||""),내용:String(r["내용"]||r["D"]||""),강사명:String(r["강사명"]||r["E"]||""),날짜,요일:String(r["요일"]||r["G"]||""),시작시간:시작,종료시간:종료,길이:calcLen(시작,종료),_src:"excel"};}).filter(r=>(r.구분||r.주제||r.내용)&&r.날짜);
        setSaving(true);
        const {data}=await supabase.from("bookings").insert(parsed.map(r=>({구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이,src:"excel"}))).select();
        setSaving(false);
        const saved=(data||[]).map(r=>({_id:r.id,구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이,_src:"excel"}));
        const merged=[...rows,...saved];const cfls=detectConflicts(merged);
        reCalc(merged);setUploadState("done");
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
    if(editIdx!=null){const target=rows[editIdx];await dbUpdate(target._id,form);reCalc(rows.map((r,i)=>i===editIdx?{...r,...form}:r));setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 수정 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);}
    else{const saved=await dbInsert({...form,_src:"manual"});reCalc([...rows,{...form,_id:saved?.id,_src:"manual"}]);setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 등록 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);}
  }
  async function handleReject(rowIdx){const target=rows[rowIdx];if(target._id)await dbDelete(target._id);reCalc(rows.filter((_,i)=>i!==rowIdx));setNotifs(prev=>prev.map(n=>(n.rowA===rowIdx||n.rowB===rowIdx)?{...n,resolved:true,unread:false}:n));}
  async function confirmCancel(){if(cancelTarget===null)return;const c=rows[cancelTarget];if(c._id)await dbDelete(c._id);reCalc(rows.filter((_,i)=>i!==cancelTarget));setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`예약 취소 — ${c?.장소||""}`,desc:`${c?.날짜} ${c?.시작시간}~${c?.종료시간} 취소됨`,time:"방금",unread:true},...prev]);setCancelTarget(null);}
  async function cancelAllConflicts(){const cfIdxs=new Set(conflicts.flatMap(c=>[c.a,c.b]));const removed=rows.filter((_,i)=>cfIdxs.has(i));for(const r of removed){if(r._id)await dbDelete(r._id);}reCalc(rows.filter((_,i)=>!cfIdxs.has(i)));setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`충돌 일괄 취소 — ${removed.length}건`,desc:`전체 제거됨`,time:"방금",unread:true},...prev]);}
  async function resetAllData(){if(!window.confirm("저장된 모든 예약 데이터를 초기화하시겠습니까?"))return;await supabase.from("bookings").delete().neq("id",0);reCalc([]);setNotifs([{id:0,type:"info",title:"초기화 완료",desc:"모든 데이터가 삭제되었습니다.",time:"방금",unread:false}]);}

  const monday=addDays(getThisWeekMonday(),weekOffset*7);
  const weekLabel=`${monday.getMonth()+1}월 ${fmtShort(monday)}(월) ~ ${fmtShort(addDays(monday,4))}(금)`;
  const unread=notifs.filter(n=>n.unread).length;
  const detectedStudios=[...new Set(rows.map(r=>r.장소).filter(Boolean))];
  const studioStats=detectedStudios.map(sid=>({id:sid,color:getColor(sid),count:rows.filter(r=>r.장소===sid).length,cf:conflicts.filter(c=>c.studio===sid).length}));
  const conflictIdxSet=new Set(conflicts.flatMap(c=>[c.a,c.b]));
  // 스케줄 목록 필터 (스튜디오 + 학기 + 날짜) + 날짜순 정렬
  const filteredRows=[...rows]
    .filter(r=>filterStudio==="전체"||r.장소===filterStudio)
    .filter(r=>filterGubun==="전체"||r.구분===filterGubun)
    .filter(r=>!filterDate||r.날짜===filterDate)
    .sort((a,b)=>{
      const da=a.날짜||"",db=b.날짜||"";
      if(da!==db) return da.localeCompare(db);
      return (toMin(a.시작시간)||0)-(toMin(b.시작시간)||0);
    });
  // 대시보드 일정표 필터 (activeStudios On/Off)
  const filteredByStudio=activeStudios.size===0?rows:rows.filter(r=>activeStudios.has(r.장소));
  // 구분값 목록
  const gubunList=["전체",...new Set(rows.map(r=>r.구분).filter(Boolean))];

  return(
    <div style={{display:"flex",flexDirection:"column",width:"100vw",height:"100vh",maxWidth:"100vw",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",fontSize:21,background:"#f5f5f4",color:"#1c1c1a",overflow:"hidden",boxSizing:"border-box"}}>
      {bookingModal==="new"&&<BookingForm title="예약 등록" initial={null} onSave={f=>handleSave(f,null)} onClose={()=>setBookingModal(null)}/>}
      {typeof bookingModal==="number"&&<BookingForm title="예약 수정" initial={rows[bookingModal]} onSave={f=>handleSave(f,bookingModal)} onClose={()=>setBookingModal(null)}/>}
      {cancelTarget!==null&&<CancelModal row={rows[cancelTarget]} onClose={()=>setCancelTarget(null)} onConfirm={confirmCancel}/>}

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 18px",height:50,background:"#fff",borderBottom:"0.5px solid #e5e5e3",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8,fontWeight:700,fontSize:22}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:"#1D9E75"}}></div>
          스튜디오 방송편성 관리
          {saving&&<span style={{fontSize:16,color:"#aaa",fontWeight:400,marginLeft:4}}>저장 중...</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button style={{...btn,color:"#A32D2D",borderColor:"#F09595",fontSize:18,height:32}} onClick={resetAllData}>🗑 초기화</button>
          <button style={{...btn,height:32,position:"relative"}} onClick={()=>setTab("notifications")}>
            🔔{unread>0&&<span style={{position:"absolute",top:4,right:4,width:7,height:7,borderRadius:"50%",background:"#E24B4A",border:"1.5px solid #fff"}}></span>}
          </button>
        </div>
      </div>

      <div style={{display:"flex",flex:1,overflow:"hidden"}}>
        <nav style={{width:175,background:"#fff",borderRight:"0.5px solid #e5e5e3",padding:"10px 0 0 0",overflowY:"auto",flexShrink:0,display:"flex",flexDirection:"column"}}>
          <div style={{fontSize:15,fontWeight:600,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 10px",marginBottom:6}}>메뉴</div>
          {[{id:"dashboard",icon:"📊",label:"대시보드"},{id:"upload",icon:"📤",label:"편성표 업로드"},{id:"schedule",icon:"📅",label:"스케줄 목록"},{id:"notifications",icon:"🔔",label:`알림${unread>0?` (${unread})`:""}`}].map(item=>(
            <div key={item.id} style={{display:"flex",alignItems:"center",gap:7,padding:"6px 10px",fontSize:18,cursor:"pointer",color:tab===item.id?"#1c1c1a":"#666",fontWeight:tab===item.id?600:400,background:tab===item.id?"#f0f0ee":"transparent",marginBottom:1}} onClick={()=>setTab(item.id)}>
              <span>{item.icon}</span>{item.label}
            </div>
          ))}
          <div style={{margin:"10px 0 6px",borderTop:"0.5px solid #e5e5e3"}}></div>
          <div style={{fontSize:15,fontWeight:600,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 10px",marginBottom:6}}>스튜디오 현황</div>
          <div style={{flex:1,overflowY:"auto"}}>
            {(studioStats.length>0?studioStats:STUDIOS.map(s=>({id:s.id,color:s.color,count:null,cf:0}))).map(st=>(
              <div key={st.id} style={{display:"flex",alignItems:"center",padding:"4px 10px",fontSize:16,color:st.count!=null?"#555":"#ccc",cursor:"pointer"}} onClick={()=>{setTab("schedule");setFilterStudio(st.id);}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:st.color,marginRight:6,flexShrink:0}}></div>
                <div style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{st.id}</div>
                {st.count!=null&&<><div style={{fontSize:15,color:"#aaa",marginRight:2}}>{st.count}</div>{st.cf>0&&<span style={{...bd("red"),padding:"0 4px",fontSize:14}}>!</span>}</>}
              </div>
            ))}
          </div>
        </nav>

        <main style={{flex:1,overflow:"hidden",padding:"10px 10px 10px 10px",minWidth:0,display:"flex",flexDirection:"column"}}>
          {tab==="dashboard"&&(
            <div style={{display:"flex",flexDirection:"column",flex:1,overflow:"hidden",gap:6}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
                <div style={{fontSize:22,fontWeight:700}}>스튜디오 방송 현황</div>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <button style={{...btn,width:28,height:28,padding:0,justifyContent:"center",fontSize:18}} onClick={()=>setWeekOffset(w=>w-1)}>◀</button>
                  <span style={{fontSize:16,fontWeight:500,minWidth:140,textAlign:"center"}}>{weekLabel}</span>
                  <button style={{...btn,width:28,height:28,padding:0,justifyContent:"center",fontSize:18}} onClick={()=>setWeekOffset(w=>w+1)}>▶</button>
                  <button style={{...btnB,height:28,fontSize:15}} onClick={()=>downloadSchedule(rows)}>⬇ 다운로드</button>
                  <button style={{...btnP,height:28,fontSize:15}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>
              <div style={{flexShrink:0}}><LiveBanner rows={rows}/></div>
              <div style={{display:"flex",gap:10,flex:1,overflow:"hidden",minHeight:0}}>
                {/* 오늘 스케줄 — 방송준비 제외 */}
                <div style={{width:"30%",flexShrink:0,display:"flex",flexDirection:"column",gap:6,overflow:"hidden",minHeight:0}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div style={{fontSize:18,fontWeight:700,color:"#1D9E75"}}>📅 오늘 라이브 스케줄</div>
                    <div style={{fontSize:14,color:"#aaa"}}>{fmtShort(new Date())} ({["일","월","화","수","목","금","토"][new Date().getDay()]}요일)</div>
                  </div>
                  <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:6}}>
                    {(()=>{
                      const todayStr=fmtFull(new Date());
                      // ★ 방송준비 제외
                      const todayRows=rows.map((r,i)=>({...r,idx:i}))
                        .filter(r=>r.날짜===todayStr&&r.장소&&r.시작시간&&r.내용!=="방송준비"&&r.구분!=="방송준비")
                        .sort((a,b)=>toMin(a.시작시간)-toMin(b.시작시간));
                      if(todayRows.length===0)return(
                        <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:"#ccc",fontSize:16,background:"#fff",borderRadius:12,border:"0.5px solid #e5e5e3",gap:8}}>
                          <div style={{fontSize:36}}>📭</div><div>오늘 예약된 라이브 방송이 없습니다</div>
                          <button style={{...btnP,height:28,fontSize:14}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                        </div>
                      );
                      const grouped={};
                      todayRows.forEach(r=>{if(!grouped[r.장소])grouped[r.장소]=[];grouped[r.장소].push(r);});
                      return Object.entries(grouped).map(([studio,sRows])=>{
                        const color=getColor(studio);
                        return(
                          <div key={studio} style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:10,overflow:"hidden",flexShrink:0}}>
                            <div style={{display:"flex",alignItems:"center",gap:6,padding:"8px 12px",background:`${color}10`,borderBottom:`1px solid ${color}30`}}>
                              <span style={{width:8,height:8,borderRadius:"50%",background:color}}></span>
                              <span style={{fontSize:18,fontWeight:700,color}}>{studio}</span>
                              <span style={{marginLeft:"auto",fontSize:15,color,background:`${color}20`,padding:"1px 7px",borderRadius:20}}>{sRows.length}건</span>
                            </div>
                            <div style={{padding:"6px 8px",display:"flex",flexDirection:"column",gap:4}}>
                              {sRows.map((b,bi)=>{
                                const isCf=conflicts.some(c=>c.a===b.idx||c.b===b.idx);
                                const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                                const content=b.내용||b.주제||"";
                                const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                                return(
                                  <div key={bi} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 8px",borderRadius:6,background:isCf?"#FFF8F8":isPrep?"#F8FFF8":"#f8f8f6",border:`0.5px solid ${isCf?"#F09595":isPrep?`${color}30`:"#f0f0ee"}`,cursor:"pointer"}} onClick={()=>setBookingModal(b.idx)}>
                                    <div style={{width:3,height:32,borderRadius:2,background:isCf?"#E24B4A":color,flexShrink:0}}></div>
                                    <div style={{flex:1,minWidth:0}}>
                                      <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:1}}>
                                        {b.구분&&!isPrep&&<span style={{fontSize:14,padding:"1px 6px",borderRadius:10,background:`${color}20`,color,fontWeight:600,flexShrink:0}}>{b.구분}</span>}
                                        {isCf&&<span style={{fontSize:14,color:"#E24B4A",fontWeight:700}}>⚠ 충돌</span>}
                                      </div>
                                      <div style={{fontSize:16,fontWeight:600,color:isCf?"#E24B4A":"#222",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{isPrep?"📋 방송준비":shortContent||content}</div>
                                      {b.강사명&&<div style={{fontSize:15,color:"#888"}}>👤 {b.강사명}</div>}
                                    </div>
                                    <div style={{fontSize:15,color:"#aaa",flexShrink:0,textAlign:"right"}}>
                                      <div style={{fontWeight:600,color}}>{b.시작시간}</div><div>~{b.종료시간}</div>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>
                <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,overflow:"hidden"}}>
                  {/* 주간 현황 헤더 + 스튜디오 필터 토글 */}
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:5,gap:8,flexShrink:0,overflow:"hidden"}}>
                    <div style={{fontSize:16,fontWeight:700,flexShrink:0,whiteSpace:"nowrap"}}>📊 주간 전체 스튜디오 현황</div>
                    <div style={{display:"flex",alignItems:"center",gap:4,overflowX:"auto",flexShrink:1,paddingBottom:2}}>
                      <button
                        style={{...btn,height:22,fontSize:11,padding:"0 8px",flexShrink:0,background:activeStudios.size===0?"#1c1c1a":"#fff",color:activeStudios.size===0?"#fff":"#555",borderColor:activeStudios.size===0?"#1c1c1a":"#ccc"}}
                        onClick={()=>setActiveStudios(new Set())}>전체</button>
                      {STUDIOS.map(s=>{
                        const on=activeStudios.has(s.id);
                        return(
                          <button key={s.id}
                            style={{...btn,height:22,fontSize:11,padding:"0 7px",flexShrink:0,background:on?s.color:"#fff",color:on?"#fff":s.color,borderColor:s.color,display:"flex",alignItems:"center",gap:3}}
                            onClick={()=>toggleStudio(s.id)}>
                            <span style={{width:5,height:5,borderRadius:"50%",background:on?"#fff":s.color,flexShrink:0}}></span>
                            {s.id}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  <div style={{flex:1,overflow:"hidden",background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12}}>
                    <div style={{width:"100%",height:"100%",overflowX:"auto",overflowY:"auto"}}>
                      <WeeklyGrid rows={rows} activeStudios={activeStudios} conflicts={conflicts} monday={monday} onEdit={setBookingModal} onCancel={setCancelTarget}/>
                    </div>
                  </div>
                </div>
              </div>
              {conflicts.length>0&&(
                <div style={{flexShrink:0,border:"0.5px solid #F09595",background:"#FFF8F8",borderRadius:10,padding:"8px 12px",maxHeight:120,overflowY:"auto"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                    <div style={{fontSize:20,fontWeight:600,color:"#A32D2D"}}>⚠️ 충돌 {conflicts.length}건</div>
                    <button style={{...btnR,height:26,fontSize:16}} onClick={()=>{if(window.confirm("충돌 예약을 모두 취소하시겠습니까?"))cancelAllConflicts();}}>🗑 전체 취소</button>
                  </div>
                  {conflicts.map((cf,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",borderBottom:"0.5px solid #F09595",fontSize:16}}>
                      <span style={{color:"#791F1F"}}>🏢 {cf.studio} · {cf.date}</span>
                      <span style={{color:"#888"}}>{cf.timeA} ↔ {cf.timeB}</span>
                      <div style={{display:"flex",gap:4}}>
                        <button style={{...btn,height:24,fontSize:15,padding:"0 8px"}} onClick={()=>setBookingModal(cf.a)}>A 수정</button>
                        <button style={{...btnR,height:24,fontSize:15,padding:"0 8px"}} onClick={()=>setCancelTarget(cf.a)}>A 취소</button>
                        <button style={{...btnR,height:24,fontSize:15,padding:"0 8px"}} onClick={()=>setCancelTarget(cf.b)}>B 취소</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
          {tab==="upload"&&(
            <>
              <div style={{fontSize:24,fontWeight:700,marginBottom:14}}>편성표 업로드</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,height:"calc(100vh - 110px)",overflow:"hidden"}}>
                <div style={{...panel,marginBottom:0,display:"flex",flexDirection:"column",overflow:"hidden"}}>
                  <div style={{fontSize:21,fontWeight:600,marginBottom:6}}>📊 엑셀 일괄 업로드</div>
                  <div style={{fontSize:18,color:"#888",marginBottom:16}}>VBA로 생성한 방송편성표 엑셀을 그대로 업로드하세요</div>
                  <div style={{flex:1,border:uploadState==="done"?"2px solid #1D9E75":"2px dashed #ccc",borderRadius:14,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",background:uploadState==="done"?"#F0FBF6":"#fafaf9",gap:12,minHeight:0}}
                    onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();handleFile({target:{files:e.dataTransfer.files}});}}>
                    <div style={{fontSize:60}}>{uploadState==="done"?"✅":uploadState==="parsing"?"⏳":uploadState==="error"?"❌":"📂"}</div>
                    <div style={{fontSize:21,fontWeight:600,color:uploadState==="done"?"#1D9E75":"#555",textAlign:"center"}}>
                      {uploadState==="idle"&&"파일을 클릭하거나 드래그하세요"}
                      {uploadState==="parsing"&&"파싱 중..."}
                      {uploadState==="done"&&fileName}
                      {uploadState==="error"&&"업로드 실패"}
                    </div>
                    <div style={{fontSize:18,color:"#aaa"}}>B열(장소) 자동 인식 · .xlsx .xls 지원</div>
                    <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleFile}/>
                  </div>
                  {parseLog.length>0&&(
                    <div style={{marginTop:12,padding:"10px 14px",borderRadius:10,background:"#f8f8f6",display:"flex",flexDirection:"column",gap:4,flexShrink:0}}>
                      {parseLog.map((l,i)=><div key={i} style={{fontSize:18,lineHeight:1.9,color:l.startsWith("⚠")||l.startsWith("❌")?"#A32D2D":"#0F6E56"}}>{l}</div>)}
                    </div>
                  )}
                </div>
                <div style={{...panel,marginBottom:0,display:"flex",flexDirection:"column",overflow:"hidden"}}>
                  <div style={{fontSize:21,fontWeight:600,marginBottom:6}}>✚ 스튜디오 직접 예약</div>
                  <div style={{fontSize:18,color:"#888",marginBottom:12}}>스튜디오를 선택해 날짜·시간을 직접 입력해 예약하세요</div>
                  <div style={{flex:1,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,alignContent:"start",overflowY:"auto",minHeight:0}}>
                    {STUDIOS.map(s=>(
                      <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",border:`1px solid ${s.color}30`,borderRadius:10,background:`${s.color}08`}}>
                        <span style={{display:"flex",alignItems:"center",gap:8}}>
                          <span style={{width:10,height:10,borderRadius:"50%",background:s.color,flexShrink:0}}></span>
                          <span style={{fontSize:16,fontWeight:500,color:"#333"}}>{s.id}</span>
                        </span>
                        <button style={{...btn,height:30,fontSize:15,padding:"0 12px",background:s.color,borderColor:s.color,color:"#fff"}} onClick={()=>setBookingModal("new")}>예약</button>
                      </div>
                    ))}
                    {/* ★ 직접 입력 예약 버튼 */}
                    <div style={{gridColumn:"1 / -1",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 16px",border:"1.5px dashed #ccc",borderRadius:10,background:"#f8f8f6",marginTop:4}}>
                      <span style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:20}}>✏</span>
                        <div>
                          <div style={{fontSize:16,fontWeight:600,color:"#555"}}>스튜디오 직접 입력 예약</div>
                          <div style={{fontSize:14,color:"#aaa"}}>목록에 없는 강의장 예약</div>
                        </div>
                      </span>
                      <button style={{...btnP,height:30,fontSize:15,padding:"0 16px"}} onClick={()=>setBookingModal("new")}>예약하기</button>
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}
          {tab==="schedule"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
                <div style={{fontSize:24,fontWeight:700}}>스케줄 목록</div>
                <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                  {/* 스튜디오 필터 */}
                  <select style={{height:34,padding:"0 8px",borderRadius:8,border:"0.5px solid #ccc",fontSize:16,background:"#fff"}} value={filterStudio} onChange={e=>setFilterStudio(e.target.value)}>
                    <option>전체</option>{STUDIOS.map(s=><option key={s.id}>{s.id}</option>)}
                  </select>
                  {/* 학기(구분) 필터 */}
                  <select style={{height:34,padding:"0 8px",borderRadius:8,border:"0.5px solid #ccc",fontSize:16,background:"#fff"}} value={filterGubun} onChange={e=>setFilterGubun(e.target.value)}>
                    {gubunList.map(g=><option key={g}>{g}</option>)}
                  </select>
                  {/* 날짜 필터 */}
                  <input type="date" style={{height:34,padding:"0 8px",borderRadius:8,border:"0.5px solid #ccc",fontSize:16,background:"#fff",fontFamily:"inherit"}} value={filterDate} onChange={e=>setFilterDate(e.target.value)}/>
                  {(filterStudio!=="전체"||filterGubun!=="전체"||filterDate)&&(
                    <button style={{...btn,height:34,fontSize:15,color:"#888"}} onClick={()=>{setFilterStudio("전체");setFilterGubun("전체");setFilterDate("");}}>✕ 초기화</button>
                  )}
                  <span style={{fontSize:18,color:"#888"}}>{filteredRows.length}건</span>
                  <button style={{...btnB,height:34,fontSize:18}} onClick={()=>downloadSchedule(filteredRows)}>⬇ 다운로드</button>
                  <button style={{...btnP,height:34,fontSize:18}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>
              <div style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,overflow:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
                  <thead><tr style={{background:"#f8f8f6",borderBottom:"0.5px solid #e5e5e3"}}>{["날짜","요일","구분","장소","내용","강사명","시작","종료","길이","출처","상태","수정","취소"].map(h=><th key={h} style={{padding:"7px 9px",fontSize:15,fontWeight:500,color:"#888",textAlign:"left",borderRight:"0.5px solid #e5e5e3",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                  <tbody>
                    {filteredRows.map((row,i)=>{
                      const realIdx=rows.indexOf(row);
                      const isCf=conflictIdxSet.has(realIdx);
                      const noSt=!row.장소;
                      const isPrep=row.내용==="방송준비"||row.구분==="방송준비";
                      const rowBg=isCf?"#FFF8F8":isPrep?"#f5f5f4":noSt?"#FFFBF0":"#fff";
                      const tdStyle={padding:"5px 9px",fontSize:16,borderRight:"0.5px solid #f0f0ee"};
                      const td=(children,extra={})=><td style={{...tdStyle,color:isPrep?"#aaa":undefined,...extra}}>{children}</td>;
                      return(
                        <tr key={i} style={{borderBottom:"0.5px solid #f0f0ee",background:rowBg}}>
                          {td(row.날짜,{whiteSpace:"nowrap"})}{td(row.요일)}
                          {td(<span style={{color:isPrep?"#aaa":"inherit",fontWeight:isPrep?400:500}}>{row.구분}</span>)}
                          {td(row.장소?<span style={{display:"inline-flex",alignItems:"center",gap:4}}><span style={{width:6,height:6,borderRadius:"50%",background:isPrep?"#ccc":getColor(row.장소),flexShrink:0}}></span><span style={{color:isPrep?"#aaa":getColor(row.장소)}}>{row.장소}</span></span>:<span style={bd("amber")}>미입력</span>)}
                          {td(row.내용||row.주제,{maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"})}
                          {td(row.강사명,{whiteSpace:"nowrap"})}{td(row.시작시간)}{td(row.종료시간)}{td(row.길이)}
                          {td(<span style={bd(row._src==="manual"?"blue":"gray")}>{row._src==="manual"?"직접":"엑셀"}</span>)}
                          {td(isCf?<span style={bd("red")}>충돌</span>:isPrep?<span style={bd("gray")}>방송준비</span>:noSt?<span style={bd("amber")}>장소미정</span>:<span style={bd("green")}>확정</span>)}
                          <td style={{padding:"5px 9px",borderRight:"0.5px solid #f0f0ee"}}><button style={{...btn,height:24,padding:"0 8px",fontSize:15,color:"#185FA5",borderColor:"#B5D4F4"}} onClick={()=>setBookingModal(realIdx)}>✏</button></td>
                          <td style={{padding:"5px 9px"}}><button style={{...btnR,height:24,padding:"0 8px",fontSize:15}} onClick={()=>setCancelTarget(realIdx)}>✕</button></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          )}
          {tab==="notifications"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <div style={{fontSize:24,fontWeight:700}}>알림</div>
                <button style={{...btn,height:30,fontSize:18}} onClick={()=>setNotifs(notifs.map(n=>({...n,unread:false})))}>모두 읽음</button>
              </div>
              {notifs.map(n=>(
                <div key={n.id} style={{background:n.unread?"#FAFFF9":"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,padding:"11px 13px",display:"flex",gap:10,marginBottom:8,opacity:n.resolved?0.55:1}}>
                  <div style={{width:30,height:30,borderRadius:8,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:21,background:n.type==="conflict"?"#FCEBEB":n.type==="cancel"?"#FFF0F0":n.type==="manual"?"#E6F1FB":"#E1F5EE",color:n.type==="conflict"?"#A32D2D":n.type==="cancel"?"#A32D2D":n.type==="manual"?"#185FA5":"#0F6E56"}}>
                    {n.type==="conflict"?"⚠️":n.type==="cancel"?"🚫":n.type==="manual"?"✏️":"📤"}
                  </div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:18,fontWeight:600,marginBottom:2}}>{n.title}</div>
                    <div style={{fontSize:16,color:"#888",lineHeight:1.6}}>{n.desc}</div>
                    {n.resolved&&<div style={{fontSize:16,color:"#0F6E56",marginTop:4}}>✓ 처리 완료</div>}
                    {n.type==="conflict"&&!n.resolved&&(
                      <div style={{marginTop:6,display:"flex",gap:5,flexWrap:"wrap"}}>
                        <button style={{...btn,height:26,fontSize:16}} onClick={()=>setBookingModal(n.rowA)}>A 수정</button>
                        <button style={{...btnR,height:26,fontSize:16}} onClick={()=>handleReject(n.rowA)}>A 반려</button>
                        <button style={{...btnR,height:26,fontSize:16}} onClick={()=>handleReject(n.rowB)}>B 반려</button>
                      </div>
                    )}
                    {n.type==="upload"&&(
                      <div style={{marginTop:6,display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{...bd("green"),fontSize:16}}>✓ DB 저장 완료</span>
                        <button style={{...btn,height:24,fontSize:16}} onClick={()=>setNotifs(prev=>prev.map(x=>x.id===n.id?{...x,unread:false}:x))}>확인</button>
                      </div>
                    )}
                  </div>
                  <div style={{fontSize:15,color:"#aaa",flexShrink:0,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
                    {n.unread&&!n.resolved&&<span style={{...bd("red"),padding:"1px 5px",fontSize:14}}>new</span>}
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
