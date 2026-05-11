import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const STUDIOS = [
  { id:"스튜디오(지하)", color:"#1D9E75" },
  { id:"17층1호",        color:"#378ADD" },
  { id:"17층2호",        color:"#7F77DD" },
  { id:"17층3호",        color:"#D4537E" },
  { id:"온택트룸2",      color:"#D85A30" },
  { id:"온택트룸4",      color:"#EF9F27" },
  { id:"온택트룸6",      color:"#639922" },
  { id:"18층 대강당",    color:"#C2410C" },
  { id:"대전",           color:"#888780" },
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

const DAYS = ["월","화","수","목","금"];
const HOURS = Array.from({length:12},(_,i)=>`${String(i+7).padStart(2,"0")}:00`);

// 샘플 데이터
const SAMPLE_ROWS = [
  {_id:1,구분:"1학기",장소:"스튜디오(지하)",주제:"코딩 Live Django트랙",내용:"방송준비",강사명:"김준호 강사",날짜:"2026-05-11",요일:"월",시작시간:"08:00",종료시간:"09:00",길이:"01:00",_src:"excel"},
  {_id:2,구분:"1학기",장소:"스튜디오(지하)",주제:"코딩 Live Django트랙",내용:"Django : DRF 1",강사명:"김준호 강사",날짜:"2026-05-11",요일:"월",시작시간:"09:00",종료시간:"16:00",길이:"07:00",_src:"excel"},
  {_id:3,구분:"1학기",장소:"17층1호",주제:"코딩 Live React트랙",내용:"방송준비",강사명:"이수민 강사",날짜:"2026-05-11",요일:"월",시작시간:"08:00",종료시간:"09:00",길이:"01:00",_src:"excel"},
  {_id:4,구분:"1학기",장소:"17층1호",주제:"코딩 Live React트랙",내용:"React : Hooks 기초",강사명:"이수민 강사",날짜:"2026-05-11",요일:"월",시작시간:"09:00",종료시간:"11:00",길이:"02:00",_src:"excel"},
  {_id:5,구분:"2학기",장소:"온택트룸2",주제:"코딩 Live Python트랙",내용:"방송준비",강사명:"박철수 강사",날짜:"2026-05-12",요일:"화",시작시간:"13:00",종료시간:"14:00",길이:"01:00",_src:"excel"},
  {_id:6,구분:"2학기",장소:"온택트룸2",주제:"코딩 Live Python트랙",내용:"Python : 알고리즘",강사명:"박철수 강사",날짜:"2026-05-12",요일:"화",시작시간:"14:00",종료시간:"16:00",길이:"02:00",_src:"excel"},
  {_id:7,구분:"취업",장소:"18층 대강당",주제:"특강",내용:"방송준비",강사명:"최민지 강사",날짜:"2026-05-13",요일:"수",시작시간:"09:00",종료시간:"10:00",길이:"01:00",_src:"manual"},
  {_id:8,구분:"취업",장소:"18층 대강당",주제:"특강",내용:"취업 특강 : 포트폴리오",강사명:"최민지 강사",날짜:"2026-05-13",요일:"수",시작시간:"10:00",종료시간:"12:00",길이:"02:00",_src:"manual"},
  {_id:9,구분:"1학기",장소:"대전",주제:"코딩 Live Java트랙",내용:"Java : Spring Boot",강사명:"홍길동 강사",날짜:"2026-05-14",요일:"목",시작시간:"09:00",종료시간:"18:00",길이:"09:00",_src:"excel"},
  {_id:11,구분:"1학기",장소:"스튜디오(지하)",주제:"코딩 Live Django트랙",내용:"방송준비",강사명:"김준호 강사",날짜:new Date().toISOString().slice(0,10),요일:["일","월","화","수","목","금","토"][new Date().getDay()],시작시간:"08:00",종료시간:"09:00",길이:"01:00",_src:"excel"},
  {_id:12,구분:"1학기",장소:"스튜디오(지하)",주제:"코딩 Live Django트랙",내용:"Django : DRF 실습",강사명:"김준호 강사",날짜:new Date().toISOString().slice(0,10),요일:["일","월","화","수","목","금","토"][new Date().getDay()],시작시간:"09:00",종료시간:"16:00",길이:"07:00",_src:"excel"},
  {_id:13,구분:"취업",장소:"18층 대강당",주제:"특강",내용:"포트폴리오 특강",강사명:"최민지 강사",날짜:new Date().toISOString().slice(0,10),요일:["일","월","화","수","목","금","토"][new Date().getDay()],시작시간:"10:00",종료시간:"12:00",길이:"02:00",_src:"manual"},
  {_id:14,구분:"2학기",장소:"온택트룸4",주제:"코딩 Live Spring트랙",내용:"Spring : JPA",강사명:"박철수 강사",날짜:new Date().toISOString().slice(0,10),요일:["일","월","화","수","목","금","토"][new Date().getDay()],시작시간:"13:00",종료시간:"16:00",길이:"03:00",_src:"excel"},
];

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

function downloadSchedule(rows){
  if(rows.length===0){alert("다운로드할 데이터가 없습니다.");return;}
  const data=rows.map(r=>({구분:r.구분,장소:r.장소,주제:r.주제,내용:r.내용,강사명:r.강사명,날짜:r.날짜,요일:r.요일,시작시간:r.시작시간,종료시간:r.종료시간,길이:r.길이}));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"방송편성표");
  ws['!cols']=[{wch:8},{wch:14},{wch:24},{wch:30},{wch:14},{wch:12},{wch:6},{wch:10},{wch:10},{wch:8}];
  XLSX.writeFile(wb,`방송편성표_${new Date().toISOString().slice(0,10)}.xlsx`);
}

const bd=(c)=>{const m={red:{bg:"#FCEBEB",fg:"#A32D2D"},green:{bg:"#E1F5EE",fg:"#0F6E56"},blue:{bg:"#E6F1FB",fg:"#185FA5"},amber:{bg:"#FAEEDA",fg:"#854F0B"},gray:{bg:"#f0f0ee",fg:"#666"}};const x=m[c]||m.gray;return{display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:500,background:x.bg,color:x.fg};};
const btn={display:"inline-flex",alignItems:"center",gap:5,padding:"0 14px",height:34,borderRadius:8,border:"0.5px solid #ccc",fontSize:13,cursor:"pointer",background:"#fff",color:"#333",fontFamily:"inherit"};
const btnP={...btn,background:"#1D9E75",borderColor:"#1D9E75",color:"#fff"};
const btnR={...btn,background:"#FCEBEB",borderColor:"#F09595",color:"#A32D2D"};
const btnB={...btn,background:"#E6F1FB",borderColor:"#B5D4F4",color:"#185FA5"};
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
    onSave({...form,길이:len});onClose();
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"#fff",borderRadius:16,width:500,maxHeight:"90vh",overflowY:"auto",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 20px",borderBottom:"0.5px solid #e5e5e3",position:"sticky",top:0,background:"#fff"}}>
          <div style={{fontSize:16,fontWeight:600}}>{title}</div>
          <button style={{...btn,height:28,padding:"0 10px"}} onClick={onClose}>✕</button>
        </div>
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",gap:12}}>
          <div>
            <span style={lbl}>스튜디오 *</span>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6}}>
              {STUDIOS.map(s=>(
                <label key={s.id} style={{display:"flex",alignItems:"center",gap:7,padding:"7px 10px",border:`1.5px solid ${form.장소===s.id?s.color:"#e5e5e3"}`,borderRadius:8,cursor:"pointer",background:form.장소===s.id?`${s.color}10`:"#fff",fontSize:12}}>
                  <input type="radio" name="bst" style={{display:"none"}} checked={form.장소===s.id} onChange={()=>set("장소",s.id)}/>
                  <span style={{width:9,height:9,borderRadius:"50%",background:s.color,flexShrink:0}}></span>
                  {s.id}{form.장소===s.id&&<span style={{marginLeft:"auto",color:s.color,fontWeight:700}}>✓</span>}
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
      <div style={{background:"#fff",borderRadius:16,width:380,padding:"24px",boxShadow:"0 12px 40px rgba(0,0,0,0.2)"}}>
        <div style={{fontSize:16,fontWeight:600,marginBottom:12}}>예약을 취소하시겠습니까?</div>
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
    <div style={{position:"fixed",top:Math.min(pos.y+10,window.innerHeight-340),left:Math.min(pos.x+10,window.innerWidth-290),zIndex:3000,background:"#fff",borderRadius:12,border:"0.5px solid #e5e5e3",boxShadow:"0 8px 28px rgba(0,0,0,0.18)",width:270,padding:"14px 16px"}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{display:"flex",alignItems:"center",gap:6,fontWeight:600,fontSize:13,color}}>
          <span style={{width:10,height:10,borderRadius:"50%",background:color}}></span>{row.장소}
        </span>
        <button style={{background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}} onClick={onClose}>✕</button>
      </div>
      <div style={{fontSize:12,lineHeight:2,color:"#444",borderBottom:"0.5px solid #f0f0ee",paddingBottom:10,marginBottom:10}}>
        {row._conflict&&<div style={{color:"#E24B4A",fontWeight:600,marginBottom:4}}>⚠ 충돌</div>}
        <div><b>구분:</b> {row.구분}</div>
        <div><b>내용:</b> {row.내용||row.주제||"-"}</div>
        {row.강사명&&<div><b>강사:</b> {row.강사명}</div>}
        <div><b>날짜:</b> {row.날짜} {row.요일&&`(${row.요일})`}</div>
        <div><b>시간:</b> {row.시작시간} ~ {row.종료시간}</div>
        {row.길이&&<div><b>길이:</b> {row.길이}</div>}
      </div>
      <div style={{display:"flex",gap:6}}>
        <button style={{...btn,flex:1,height:30,fontSize:12,justifyContent:"center"}} onClick={()=>{onEdit(idx);onClose();}}>✏ 수정</button>
        <button style={{...btnR,flex:1,height:30,fontSize:12,justifyContent:"center"}} onClick={()=>{onCancel(idx);onClose();}}>✕ 취소</button>
      </div>
    </div>
  );
}

function WeeklyGrid({rows,conflicts,monday,onEdit,onCancel}){
  const [popup,setPopup]=useState(null);
  const dayDates=DAYS.map((_,i)=>fmtFull(addDays(monday,i)));
  const conflictIdxs=new Set();
  conflicts.forEach(c=>{conflictIdxs.add(c.a);conflictIdxs.add(c.b);});

  // 특정 시간대에 걸치는 블록 찾기 (시작~종료 사이에 해당 hour 포함)
  function getBlocksInHour(dayDate, hourStr){
    const hMin=toMin(hourStr);
    return rows
      .map((r,i)=>({...r,idx:i}))
      .filter(r=>{
        if(r.날짜!==dayDate||!r.시작시간||!r.장소) return false;
        const s=toMin(r.시작시간), e=toMin(r.종료시간);
        if(s===null||e===null) return false;
        return s<hMin+60 && e>hMin; // 이 시간대에 겹치는 블록
      })
      .sort((a,b)=>toMin(a.시작시간)-toMin(b.시작시간));
  }

  return(
    <div style={{position:"relative",width:"100%"}} onClick={()=>setPopup(null)}>
      {popup&&<BlockPopup row={popup.row} idx={popup.idx} pos={{x:popup.x,y:popup.y}} onClose={()=>setPopup(null)} onEdit={onEdit} onCancel={onCancel}/>}
      <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
        <thead>
          <tr style={{position:"sticky",top:0,zIndex:2}}>
            <th style={{width:58,padding:"8px 6px",background:"#f8f8f6",borderBottom:"2px solid #e5e5e3",borderRight:"0.5px solid #e5e5e3",fontSize:11,fontWeight:600,color:"#888",textAlign:"center"}}>시간</th>
            {DAYS.map((day,di)=>{
              const isToday=fmtFull(new Date())===dayDates[di];
              const dayCount=rows.filter(r=>r.날짜===dayDates[di]&&r.장소).length;
              return(
                <th key={day} style={{padding:"8px 6px",background:isToday?"#F0FBF6":"#f8f8f6",borderBottom:"2px solid #e5e5e3",borderRight:"0.5px solid #e5e5e3",fontSize:12,fontWeight:700,color:isToday?"#1D9E75":"#222",textAlign:"center"}}>
                  <div>{day}요일</div>
                  <div style={{fontSize:10,fontWeight:400,color:isToday?"#1D9E75":"#aaa"}}>{fmtShort(addDays(monday,di))}</div>
                  {dayCount>0&&<div style={{fontSize:9,color:isToday?"#1D9E75":"#aaa",marginTop:1}}>{dayCount}건</div>}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {HOURS.map((hour,hi)=>(
            <tr key={hour} style={{background:hi%2===0?"#fff":"#fafafa"}}>
              {/* 시간 셀 */}
              <td style={{padding:"6px 4px",borderBottom:"0.5px solid #f0f0ee",borderRight:"0.5px solid #e5e5e3",textAlign:"center",verticalAlign:"top",background:hi%2===0?"#fafaf9":"#f5f5f4"}}>
                <span style={{fontSize:11,fontWeight:600,color:"#aaa"}}>{hour}</span>
              </td>
              {/* 요일별 셀 */}
              {DAYS.map((day,di)=>{
                const isToday=fmtFull(new Date())===dayDates[di];
                const blocks=getBlocksInHour(dayDates[di], hour);
                // 이 시간대의 시작 블록만 표시 (중간에 걸친 건 첫 시간에만)
                const startBlocks=blocks.filter(b=>{
                  const s=toMin(b.시작시간);
                  const hMin=toMin(hour);
                  return s>=hMin && s<hMin+60;
                });
                // 이전 시간에서 이어지는 블록 (시작이 이 시간대 이전)
                const continueBlocks=blocks.filter(b=>{
                  const s=toMin(b.시작시간);
                  const hMin=toMin(hour);
                  return s<hMin;
                });
                return(
                  <td key={day} style={{padding:"3px 4px",borderBottom:"0.5px solid #f0f0ee",borderRight:"0.5px solid #e5e5e3",verticalAlign:"top",minHeight:36,background:isToday?"rgba(29,158,117,0.02)":"transparent"}}>
                    {startBlocks.length===0&&continueBlocks.length===0&&(
                      <div style={{height:32,display:"flex",alignItems:"center",justifyContent:"center",color:"#f0f0ee",fontSize:10}}>·</div>
                    )}
                    {/* 이어지는 블록 (연한 표시) */}
                    {continueBlocks.map((b,bi)=>{
                      const color=getColor(b.장소);
                      return(
                        <div key={`c${bi}`} style={{borderLeft:`3px solid ${color}`,background:`${color}08`,borderRadius:"0 4px 4px 0",padding:"2px 5px",marginBottom:2,cursor:"pointer",opacity:0.7}}
                          onClick={e=>{e.stopPropagation();setPopup({row:{...b,_conflict:conflictIdxs.has(b.idx)},idx:b.idx,x:e.clientX,y:e.clientY});}}>
                          <div style={{fontSize:9,color,fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>↑ {b.장소}</div>
                        </div>
                      );
                    })}
                    {/* 이 시간에 시작하는 블록 */}
                    {startBlocks.map((b,bi)=>{
                      const isCf=conflictIdxs.has(b.idx);
                      const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                      const color=isCf?"#E24B4A":getColor(b.장소);
                      const content=b.내용||b.주제||"";
                      const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                      const dur=calcLen(b.시작시간,b.종료시간);
                      return(
                        <div key={bi}
                          style={{background:`${color}12`,border:`1.5px solid ${color}`,borderLeft:`3px solid ${color}`,borderRadius:5,padding:"4px 6px",marginBottom:3,cursor:"pointer",transition:"box-shadow 0.15s"}}
                          onClick={e=>{e.stopPropagation();setPopup({row:{...b,_conflict:isCf},idx:b.idx,x:e.clientX,y:e.clientY});}}
                          onMouseEnter={e=>e.currentTarget.style.boxShadow=`0 2px 8px ${color}40`}
                          onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}
                        >
                          {isCf&&<div style={{fontSize:8,color:"#E24B4A",fontWeight:700}}>⚠ 충돌</div>}
                          <div style={{fontSize:9,fontWeight:700,color,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                            {b.장소}
                          </div>
                          <div style={{fontSize:9,fontWeight:600,color:"#333",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                            {isPrep?"📋 방송준비":shortContent||b.주제||b.내용}
                          </div>
                          {b.강사명&&<div style={{fontSize:8,color:"#888",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>👤 {b.강사명}</div>}
                          <div style={{fontSize:8,color:"#aaa"}}>{b.시작시간}~{b.종료시간}{dur?` (${dur})`:""}</div>
                          <div style={{display:"flex",gap:2,marginTop:3}}>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:`0.5px solid ${color}`,borderRadius:3,fontSize:8,cursor:"pointer",padding:"1px 0",color,fontWeight:700}} onClick={e=>{e.stopPropagation();onEdit(b.idx);}}>✏ 수정</button>
                            <button style={{flex:1,background:"rgba(255,255,255,0.9)",border:"0.5px solid #F09595",borderRadius:3,fontSize:8,cursor:"pointer",padding:"1px 0",color:"#A32D2D",fontWeight:700}} onClick={e=>{e.stopPropagation();onCancel(b.idx);}}>✕ 취소</button>
                          </div>
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

export default function App(){
  const [tab,setTab]=useState("dashboard");
  const [rows,setRows]=useState(SAMPLE_ROWS);
  const [conflicts,setConflicts]=useState(detectConflicts(SAMPLE_ROWS));
  const [notifs,setNotifs]=useState([{id:0,type:"info",title:"미리보기 모드",desc:"샘플 데이터로 동작 중입니다. 실제 서비스는 Vercel 배포 버전을 사용하세요.",time:"",unread:false}]);
  const [uploadState,setUploadState]=useState("idle");
  const [parseLog,setParseLog]=useState([]);
  const [weekOffset,setWeekOffset]=useState(0);
  const [bookingModal,setBookingModal]=useState(null);
  const [cancelTarget,setCancelTarget]=useState(null);
  const [filterStudio,setFilterStudio]=useState("전체");
  const [fileName,setFileName]=useState("");
  const fileRef=useRef();

  function reCalc(newRows){setRows(newRows);setConflicts(detectConflicts(newRows));}

  function handleFile(e){
    const file=e.target.files[0];
    if(!file) return;
    setFileName(file.name);setUploadState("parsing");
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try{
        const wb=XLSX.read(evt.target.result,{type:"binary",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
        const parsed=raw.map((r,i)=>{
          const 날짜=excelDateToStr(r["날짜"]||r["F"]||"");
          const 시작=excelTimeToStr(r["시작시간"]||r["H"]||"");
          const 종료=excelTimeToStr(r["종료시간"]||r["I"]||"");
          return{_id:Date.now()+i,구분:String(r["구분"]||r["A"]||""),장소:normStudio(r["장소"]||r["B"]||""),주제:String(r["주제"]||r["C"]||""),내용:String(r["내용"]||r["D"]||""),강사명:String(r["강사명"]||r["E"]||""),날짜,요일:String(r["요일"]||r["G"]||""),시작시간:시작,종료시간:종료,길이:calcLen(시작,종료),_src:"excel"};
        }).filter(r=>(r.구분||r.주제||r.내용)&&r.날짜);
        const merged=[...rows,...parsed];
        const cfls=detectConflicts(merged);
        reCalc(merged);setUploadState("done");
        const studios=[...new Set(parsed.map(r=>r.장소).filter(Boolean))];
        setParseLog([`✅ ${parsed.length}행 파싱 완료`,`🏢 스튜디오: ${studios.join(", ")||"없음"}`,cfls.length>0?`⚠️ 충돌 ${cfls.length}건`:`✓ 충돌 없음`]);
        setNotifs(prev=>[{id:Date.now(),type:"upload",title:`업로드 완료 — ${file.name}`,desc:`${parsed.length}건 반영됨`,time:"방금",unread:true},...prev]);
      }catch(err){setUploadState("error");setParseLog(["❌ 파싱 실패: "+err.message]);}
    };
    reader.readAsBinaryString(file);
  }

  function handleSave(form,editIdx){
    if(editIdx!=null){
      reCalc(rows.map((r,i)=>i===editIdx?{...r,...form}:r));
      setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 수정 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);
    } else {
      reCalc([...rows,{...form,_id:Date.now(),_src:"manual"}]);
      setNotifs(prev=>[{id:Date.now(),type:"manual",title:`예약 등록 — ${form.장소}`,desc:`${form.날짜} ${form.시작시간}~${form.종료시간}`,time:"방금",unread:true},...prev]);
    }
  }

  function handleReject(rowIdx){
    reCalc(rows.filter((_,i)=>i!==rowIdx));
    setNotifs(prev=>prev.map(n=>(n.rowA===rowIdx||n.rowB===rowIdx)?{...n,resolved:true,unread:false}:n));
  }

  function confirmCancel(){
    if(cancelTarget===null) return;
    const c=rows[cancelTarget];
    reCalc(rows.filter((_,i)=>i!==cancelTarget));
    setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`예약 취소 — ${c?.장소||""}`,desc:`${c?.날짜} ${c?.시작시간}~${c?.종료시간} 취소됨`,time:"방금",unread:true},...prev]);
    setCancelTarget(null);
  }

  function cancelAllConflicts(){
    const cfIdxs=new Set(conflicts.flatMap(c=>[c.a,c.b]));
    const removed=rows.filter((_,i)=>cfIdxs.has(i));
    reCalc(rows.filter((_,i)=>!cfIdxs.has(i)));
    setNotifs(prev=>[{id:Date.now(),type:"cancel",title:`충돌 일괄 취소 — ${removed.length}건`,desc:`전체 제거됨`,time:"방금",unread:true},...prev]);
  }

  const monday=addDays(getThisWeekMonday(),weekOffset*7);
  const weekLabel=`${monday.getMonth()+1}월 ${fmtShort(monday)}(월) ~ ${fmtShort(addDays(monday,4))}(금)`;
  const unread=notifs.filter(n=>n.unread).length;
  const detectedStudios=[...new Set(rows.map(r=>r.장소).filter(Boolean))];
  const studioStats=detectedStudios.map(sid=>({id:sid,color:getColor(sid),count:rows.filter(r=>r.장소===sid).length,cf:conflicts.filter(c=>c.studio===sid).length}));
  const conflictIdxSet=new Set(conflicts.flatMap(c=>[c.a,c.b]));
  const filteredRows=filterStudio==="전체"?rows:rows.filter(r=>r.장소===filterStudio);

  return(
    <div style={{display:"flex",flexDirection:"column",width:"100%",height:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",fontSize:14,background:"#f5f5f4",color:"#1c1c1a",overflow:"hidden"}}>
      {bookingModal==="new"&&<BookingForm title="예약 등록" initial={null} onSave={f=>handleSave(f,null)} onClose={()=>setBookingModal(null)}/>}
      {typeof bookingModal==="number"&&<BookingForm title="예약 수정" initial={rows[bookingModal]} onSave={f=>handleSave(f,bookingModal)} onClose={()=>setBookingModal(null)}/>}
      {cancelTarget!==null&&<CancelModal row={rows[cancelTarget]} onClose={()=>setCancelTarget(null)} onConfirm={confirmCancel}/>}

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 16px",height:50,background:"#fff",borderBottom:"0.5px solid #e5e5e3",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8,fontWeight:700,fontSize:14}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:"#1D9E75"}}></div>
          스튜디오 방송편성 관리
          <span style={{fontSize:10,color:"#aaa",fontWeight:400,background:"#f0f0ee",padding:"2px 7px",borderRadius:10}}>미리보기</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <button style={{...btn,height:30,fontSize:12,position:"relative"}} onClick={()=>setTab("notifications")}>
            🔔{unread>0&&<span style={{position:"absolute",top:3,right:3,width:7,height:7,borderRadius:"50%",background:"#E24B4A",border:"1.5px solid #fff"}}></span>}
          </button>
        </div>
      </div>

      <div style={{display:"flex",flex:1,overflow:"hidden"}}>
        {/* 사이드바 — 좌우 패딩 제거해서 빈공간 없앰 */}
        <nav style={{width:175,background:"#fff",borderRight:"0.5px solid #e5e5e3",padding:"12px 0",overflowY:"auto",flexShrink:0}}>
          <div style={{fontSize:10,fontWeight:600,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 10px",marginBottom:6}}>메뉴</div>
          {[{id:"dashboard",icon:"📊",label:"대시보드"},{id:"upload",icon:"📤",label:"편성표 업로드"},{id:"schedule",icon:"📅",label:"스케줄 목록"},{id:"notifications",icon:"🔔",label:`알림${unread>0?` (${unread})`:""}`}].map(item=>(
            <div key={item.id} style={{display:"flex",alignItems:"center",gap:7,padding:"6px 10px",fontSize:12,cursor:"pointer",color:tab===item.id?"#1c1c1a":"#666",fontWeight:tab===item.id?600:400,background:tab===item.id?"#f0f0ee":"transparent",marginBottom:1}} onClick={()=>setTab(item.id)}>
              <span>{item.icon}</span>{item.label}
            </div>
          ))}
          <div style={{margin:"10px 0 6px",borderTop:"0.5px solid #e5e5e3"}}></div>
          <div style={{fontSize:10,fontWeight:600,color:"#999",textTransform:"uppercase",letterSpacing:"0.6px",padding:"0 10px",marginBottom:6}}>스튜디오 현황</div>
          {(studioStats.length>0?studioStats:STUDIOS.map(s=>({id:s.id,color:s.color,count:null,cf:0}))).map(st=>(
            <div key={st.id} style={{display:"flex",alignItems:"center",padding:"4px 10px",fontSize:11,color:st.count!=null?"#555":"#ccc",cursor:"pointer"}} onClick={()=>{setTab("schedule");setFilterStudio(st.id);}}>
              <div style={{width:7,height:7,borderRadius:"50%",background:st.color,marginRight:6,flexShrink:0}}></div>
              <div style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{st.id}</div>
              {st.count!=null&&<><div style={{fontSize:10,color:"#aaa",marginRight:2}}>{st.count}</div>{st.cf>0&&<span style={{...bd("red"),padding:"0 4px",fontSize:9}}>!</span>}</>}
            </div>
          ))}
        </nav>

        <main style={{flex:1,overflowY:"auto",padding:14,minWidth:0}}>

          {tab==="dashboard"&&(
            <>
              {/* 상단 헤더 */}
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
                <div style={{fontSize:16,fontWeight:700}}>스튜디오 방송 현황</div>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <button style={{...btn,width:28,height:28,padding:0,justifyContent:"center",fontSize:12}} onClick={()=>setWeekOffset(w=>w-1)}>◀</button>
                  <span style={{fontSize:12,fontWeight:500,minWidth:150,textAlign:"center"}}>{weekLabel}</span>
                  <button style={{...btn,width:28,height:28,padding:0,justifyContent:"center",fontSize:12}} onClick={()=>setWeekOffset(w=>w+1)}>▶</button>
                  <button style={{...btnB,height:28,fontSize:11}} onClick={()=>downloadSchedule(rows)}>⬇ 다운로드</button>
                  <button style={{...btnP,height:28,fontSize:11}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>

              {/* 현재 진행 중인 라이브 롤링 배너 */}
              {(()=>{
                const nowMin=new Date().getHours()*60+new Date().getMinutes();
                const live=rows.filter(r=>{
                  const s=toMin(r.시작시간),e=toMin(r.종료시간);
                  return s!==null&&e!==null&&s<=nowMin&&e>nowMin&&r.장소;
                });
                const [rollIdx,setRollIdx]=useState(0);
                useEffect(()=>{
                  if(live.length<=1) return;
                  const t=setInterval(()=>setRollIdx(i=>(i+1)%live.length),2800);
                  return()=>clearInterval(t);
                },[live.length]);

                if(live.length===0) return(
                  <div style={{display:"flex",alignItems:"center",gap:10,background:"#f8f8f6",border:"0.5px solid #e5e5e3",borderRadius:10,padding:"10px 16px",marginBottom:10}}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:"#ccc"}}></div>
                    <span style={{fontSize:12,color:"#aaa"}}>현재 진행 중인 라이브 방송이 없습니다</span>
                    <span style={{marginLeft:"auto",fontSize:11,color:"#ccc"}}>
                      {String(new Date().getHours()).padStart(2,"0")}:{String(new Date().getMinutes()).padStart(2,"0")} 기준
                    </span>
                  </div>
                );

                const cur=live[rollIdx%live.length];
                const color=getColor(cur.장소);
                const content=cur.내용||cur.주제||"";
                const shortContent=content.includes(":")?content.split(":").pop().trim():content;

                return(
                  <div style={{display:"flex",alignItems:"center",gap:12,background:`${color}10`,border:`1.5px solid ${color}40`,borderRadius:10,padding:"10px 16px",marginBottom:10,overflow:"hidden",position:"relative"}}>
                    {/* LIVE 뱃지 */}
                    <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                      <div style={{width:8,height:8,borderRadius:"50%",background:"#E24B4A",animation:"pulse 1.2s infinite"}}></div>
                      <span style={{fontSize:11,fontWeight:700,color:"#E24B4A",letterSpacing:"0.5px"}}>LIVE</span>
                    </div>
                    <div style={{width:"1px",height:28,background:`${color}40`,flexShrink:0}}></div>
                    {/* 스튜디오 */}
                    <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                      <span style={{width:8,height:8,borderRadius:"50%",background:color}}></span>
                      <span style={{fontSize:12,fontWeight:700,color}}>{cur.장소}</span>
                    </div>
                    {/* 내용 */}
                    <div style={{flex:1,minWidth:0}}>
                      <span style={{fontSize:12,fontWeight:600,color:"#222",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",display:"block"}}>
                        {shortContent||content}
                      </span>
                      {cur.강사명&&<span style={{fontSize:11,color:"#666"}}>👤 {cur.강사명}</span>}
                    </div>
                    {/* 시간 */}
                    <div style={{flexShrink:0,textAlign:"right"}}>
                      <div style={{fontSize:11,fontWeight:600,color}}>{cur.시작시간}~{cur.종료시간}</div>
                      {cur.길이&&<div style={{fontSize:10,color:"#aaa"}}>{cur.길이}</div>}
                    </div>
                    {/* 인디케이터 */}
                    {live.length>1&&(
                      <div style={{display:"flex",gap:3,flexShrink:0,marginLeft:4}}>
                        {live.map((_,i)=>(
                          <div key={i} style={{width:5,height:5,borderRadius:"50%",background:i===rollIdx%live.length?color:"#ddd",transition:"background 0.3s"}}></div>
                        ))}
                      </div>
                    )}
                    {live.length>1&&(
                      <div style={{position:"absolute",right:16,top:6,fontSize:10,color:"#aaa"}}>{rollIdx%live.length+1}/{live.length}</div>
                    )}
                    <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}`}</style>
                  </div>
                );
              })()}

              {/* ── 메인 2분할 레이아웃 ── */}
              <div style={{display:"flex",gap:12,height:"calc(100vh - 280px)",minHeight:400}}>

                {/* 좌측: 오늘 스케줄 */}
                <div style={{width:"32%",flexShrink:0,display:"flex",flexDirection:"column",gap:8}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div style={{fontSize:13,fontWeight:700,color:"#1D9E75"}}>
                      📅 오늘 스케줄
                    </div>
                    <div style={{fontSize:11,color:"#aaa"}}>
                      {fmtShort(new Date())} ({["일","월","화","수","목","금","토"][new Date().getDay()]}요일)
                    </div>
                  </div>
                  <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:6}}>
                    {(()=>{
                      const todayStr=fmtFull(new Date());
                      const todayRows=rows
                        .map((r,i)=>({...r,idx:i}))
                        .filter(r=>r.날짜===todayStr&&r.장소&&r.시작시간)
                        .sort((a,b)=>toMin(a.시작시간)-toMin(b.시작시간));
                      if(todayRows.length===0) return(
                        <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:"#ccc",fontSize:12,background:"#fff",borderRadius:12,border:"0.5px solid #e5e5e3",gap:8}}>
                          <div style={{fontSize:28}}>📭</div>
                          <div>오늘 예약된 스케줄이 없습니다</div>
                          <button style={{...btnP,height:28,fontSize:11}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                        </div>
                      );
                      // 스튜디오별로 그룹핑
                      const grouped={};
                      todayRows.forEach(r=>{
                        if(!grouped[r.장소]) grouped[r.장소]=[];
                        grouped[r.장소].push(r);
                      });
                      return Object.entries(grouped).map(([studio,sRows])=>{
                        const color=getColor(studio);
                        return(
                          <div key={studio} style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:10,overflow:"hidden",flexShrink:0}}>
                            {/* 스튜디오 헤더 */}
                            <div style={{display:"flex",alignItems:"center",gap:6,padding:"8px 12px",background:`${color}10`,borderBottom:`1px solid ${color}30`}}>
                              <span style={{width:8,height:8,borderRadius:"50%",background:color}}></span>
                              <span style={{fontSize:12,fontWeight:700,color}}>{studio}</span>
                              <span style={{marginLeft:"auto",fontSize:10,color,background:`${color}20`,padding:"1px 7px",borderRadius:20}}>{sRows.length}건</span>
                            </div>
                            {/* 예약 목록 */}
                            <div style={{padding:"6px 8px",display:"flex",flexDirection:"column",gap:4}}>
                              {sRows.map((b,bi)=>{
                                const isCf=conflicts.some(c=>c.a===b.idx||c.b===b.idx);
                                const isPrep=b.내용==="방송준비"||b.구분==="방송준비";
                                const content=b.내용||b.주제||"";
                                const shortContent=content.includes(":")?content.split(":").pop().trim():content;
                                return(
                                  <div key={bi} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 8px",borderRadius:6,background:isCf?"#FFF8F8":isPrep?"#F8FFF8":"#f8f8f6",border:`0.5px solid ${isCf?"#F09595":isPrep?`${color}30`:"#f0f0ee"}`,cursor:"pointer"}}
                                    onClick={()=>setBookingModal(b.idx)}>
                                    <div style={{width:3,height:32,borderRadius:2,background:isCf?"#E24B4A":color,flexShrink:0}}></div>
                                    <div style={{flex:1,minWidth:0}}>
                                      <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:1}}>
                                        {b.구분&&!isPrep&&<span style={{fontSize:9,padding:"1px 6px",borderRadius:10,background:`${color}20`,color,fontWeight:600,flexShrink:0}}>{b.구분}</span>}
                                        {isCf&&<span style={{fontSize:9,color:"#E24B4A",fontWeight:700}}>⚠ 충돌</span>}
                                      </div>
                                      <div style={{fontSize:11,fontWeight:600,color:isCf?"#E24B4A":"#222",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                                        {isPrep?"📋 방송준비":shortContent||content}
                                      </div>
                                      {b.강사명&&<div style={{fontSize:10,color:"#888"}}>👤 {b.강사명}</div>}
                                    </div>
                                    <div style={{fontSize:10,color:"#aaa",flexShrink:0,textAlign:"right"}}>
                                      <div style={{fontWeight:600,color}}>{b.시작시간}</div>
                                      <div>~{b.종료시간}</div>
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

                {/* 우측: 주간 전체 현황 (스크롤) */}
                <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
                    <div style={{fontSize:13,fontWeight:700}}>📊 주간 전체 스튜디오 현황</div>
                    <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                      {STUDIOS.slice(0,6).map(s=>(
                        <span key={s.id} style={{display:"flex",alignItems:"center",gap:3,fontSize:9,color:"#555",padding:"1px 6px",borderRadius:20,background:`${s.color}15`,border:`0.5px solid ${s.color}50`}}>
                          <span style={{width:5,height:5,borderRadius:"50%",background:s.color}}></span>{s.id}
                        </span>
                      ))}
                      <span style={{fontSize:9,color:"#aaa"}}>+{STUDIOS.length-6}개</span>
                    </div>
                  </div>
                  {/* 스크롤 가능한 주간 테이블 */}
                  <div style={{flex:1,overflowY:"auto",overflowX:"auto",background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12}}>
                    <WeeklyGrid rows={rows} conflicts={conflicts} monday={monday} onEdit={setBookingModal} onCancel={setCancelTarget}/>
                  </div>
                </div>
              </div>

              {/* 충돌 섹션 */}
              {conflicts.length>0&&(
                <div style={{...panel,border:"0.5px solid #F09595",background:"#FFF8F8",marginTop:10}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                    <div style={{fontSize:13,fontWeight:600,color:"#A32D2D"}}>⚠️ 충돌 {conflicts.length}건</div>
                    <button style={{...btnR,height:26,fontSize:11}} onClick={()=>{if(window.confirm("충돌 예약을 모두 취소하시겠습니까?"))cancelAllConflicts();}}>🗑 전체 취소</button>
                  </div>
                  {conflicts.map((cf,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",borderBottom:"0.5px solid #F09595",fontSize:11}}>
                      <span style={{color:"#791F1F"}}>🏢 {cf.studio} · {cf.date}</span>
                      <span style={{color:"#888"}}>{cf.timeA} ↔ {cf.timeB}</span>
                      <div style={{display:"flex",gap:4}}>
                        <button style={{...btn,height:24,fontSize:10,padding:"0 8px"}} onClick={()=>setBookingModal(cf.a)}>A 수정</button>
                        <button style={{...btnR,height:24,fontSize:10,padding:"0 8px"}} onClick={()=>setCancelTarget(cf.a)}>A 취소</button>
                        <button style={{...btnR,height:24,fontSize:10,padding:"0 8px"}} onClick={()=>setCancelTarget(cf.b)}>B 취소</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}

          {tab==="upload"&&(
            <>
              <div style={{fontSize:16,fontWeight:700,marginBottom:14}}>편성표 업로드</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div style={panel}>
                  <div style={{fontSize:13,fontWeight:600,marginBottom:4}}>📊 엑셀 일괄 업로드</div>
                  <div style={{fontSize:11,color:"#888",marginBottom:12}}>VBA로 생성한 방송편성표 엑셀 업로드</div>
                  <div style={{border:uploadState==="done"?"1.5px solid #1D9E75":"1px dashed #ccc",borderRadius:10,padding:20,textAlign:"center",cursor:"pointer",marginBottom:8}} onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();handleFile({target:{files:e.dataTransfer.files}});}}>
                    <div style={{fontSize:24,marginBottom:6}}>{uploadState==="done"?"✅":uploadState==="parsing"?"⏳":uploadState==="error"?"❌":"📂"}</div>
                    <div style={{fontSize:12,fontWeight:500,marginBottom:2}}>{uploadState==="idle"?"파일 클릭 또는 드래그":uploadState==="parsing"?"처리 중...":uploadState==="done"?fileName:"실패"}</div>
                    <div style={{fontSize:10,color:"#aaa"}}>B열(장소) 자동 인식 · .xlsx .xls</div>
                    <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleFile}/>
                  </div>
                  {parseLog.map((l,i)=><div key={i} style={{fontSize:11,lineHeight:1.9,color:l.startsWith("⚠")||l.startsWith("❌")?"#A32D2D":"#0F6E56"}}>{l}</div>)}
                </div>
                <div style={panel}>
                  <div style={{fontSize:13,fontWeight:600,marginBottom:4}}>✚ 스튜디오 직접 예약</div>
                  <div style={{fontSize:11,color:"#888",marginBottom:12}}>날짜·시간을 직접 입력해 예약</div>
                  <div style={{display:"flex",flexDirection:"column",gap:5,maxHeight:320,overflowY:"auto"}}>
                    {STUDIOS.map(s=>(
                      <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"7px 10px",border:"0.5px solid #e5e5e3",borderRadius:8,fontSize:11}}>
                        <span style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:7,height:7,borderRadius:"50%",background:s.color}}></span>{s.id}</span>
                        <button style={{...btn,height:24,fontSize:10,padding:"0 8px"}} onClick={()=>setBookingModal("new")}>예약하기</button>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </>
          )}

          {tab==="schedule"&&(
            <>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <div style={{fontSize:16,fontWeight:700}}>스케줄 목록</div>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <select style={{height:32,padding:"0 8px",borderRadius:8,border:"0.5px solid #ccc",fontSize:11,background:"#fff"}} value={filterStudio} onChange={e=>setFilterStudio(e.target.value)}>
                    <option>전체</option>
                    {STUDIOS.map(s=><option key={s.id}>{s.id}</option>)}
                  </select>
                  <span style={{fontSize:12,color:"#888"}}>{filteredRows.length}건</span>
                  <button style={{...btnB,height:30,fontSize:12}} onClick={()=>downloadSchedule(filteredRows)}>⬇ 다운로드</button>
                  <button style={{...btnP,height:30,fontSize:12}} onClick={()=>setBookingModal("new")}>✚ 예약 등록</button>
                </div>
              </div>
              <div style={{background:"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,overflow:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",minWidth:820}}>
                  <thead>
                    <tr style={{background:"#f8f8f6",borderBottom:"0.5px solid #e5e5e3"}}>
                      {["날짜","요일","구분","장소","내용","강사명","시작","종료","길이","출처","상태","수정","취소"].map(h=>(
                        <th key={h} style={{padding:"7px 9px",fontSize:10,fontWeight:500,color:"#888",textAlign:"left",borderRight:"0.5px solid #e5e5e3",whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row,i)=>{
                      const realIdx=rows.indexOf(row);
                      const isCf=conflictIdxSet.has(realIdx);
                      const noSt=!row.장소;
                      const td=(children,extra={})=><td style={{padding:"5px 9px",fontSize:11,borderRight:"0.5px solid #f0f0ee",...extra}}>{children}</td>;
                      return(
                        <tr key={i} style={{borderBottom:"0.5px solid #f0f0ee",background:isCf?"#FFF8F8":noSt?"#FFFBF0":"#fff"}}>
                          {td(row.날짜,{whiteSpace:"nowrap"})}
                          {td(row.요일)}
                          {td(row.구분)}
                          {td(row.장소?<span style={{display:"inline-flex",alignItems:"center",gap:4}}><span style={{width:6,height:6,borderRadius:"50%",background:getColor(row.장소),flexShrink:0}}></span>{row.장소}</span>:<span style={bd("amber")}>미입력</span>)}
                          {td(row.내용||row.주제,{maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"})}
                          {td(row.강사명,{whiteSpace:"nowrap"})}
                          {td(row.시작시간)}
                          {td(row.종료시간)}
                          {td(row.길이)}
                          {td(<span style={bd(row._src==="manual"?"blue":"gray")}>{row._src==="manual"?"직접":"엑셀"}</span>)}
                          {td(isCf?<span style={bd("red")}>충돌</span>:noSt?<span style={bd("amber")}>장소미정</span>:<span style={bd("green")}>확정</span>)}
                          <td style={{padding:"5px 9px",borderRight:"0.5px solid #f0f0ee"}}><button style={{...btn,height:24,padding:"0 8px",fontSize:10,color:"#185FA5",borderColor:"#B5D4F4"}} onClick={()=>setBookingModal(realIdx)}>✏</button></td>
                          <td style={{padding:"5px 9px"}}><button style={{...btnR,height:24,padding:"0 8px",fontSize:10}} onClick={()=>setCancelTarget(realIdx)}>✕</button></td>
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
                <div style={{fontSize:16,fontWeight:700}}>알림</div>
                <button style={{...btn,height:30,fontSize:12}} onClick={()=>setNotifs(notifs.map(n=>({...n,unread:false})))}>모두 읽음</button>
              </div>
              {notifs.map(n=>(
                <div key={n.id} style={{background:n.unread?"#FAFFF9":"#fff",border:"0.5px solid #e5e5e3",borderRadius:12,padding:"11px 13px",display:"flex",gap:10,marginBottom:8,opacity:n.resolved?0.55:1}}>
                  <div style={{width:30,height:30,borderRadius:8,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,background:n.type==="conflict"?"#FCEBEB":n.type==="cancel"?"#FFF0F0":n.type==="manual"?"#E6F1FB":"#E1F5EE",color:n.type==="conflict"?"#A32D2D":n.type==="cancel"?"#A32D2D":n.type==="manual"?"#185FA5":"#0F6E56"}}>
                    {n.type==="conflict"?"⚠️":n.type==="cancel"?"🚫":n.type==="manual"?"✏️":"📤"}
                  </div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:600,marginBottom:2}}>{n.title}</div>
                    <div style={{fontSize:11,color:"#888",lineHeight:1.6}}>{n.desc}</div>
                    {n.type==="conflict"&&!n.resolved&&(
                      <div style={{marginTop:6,display:"flex",gap:5,flexWrap:"wrap"}}>
                        <button style={{...btn,height:26,fontSize:11}} onClick={()=>setBookingModal(n.rowA)}>A 수정</button>
                        <button style={{...btnR,height:26,fontSize:11}} onClick={()=>handleReject(n.rowA)}>A 반려</button>
                        <button style={{...btnR,height:26,fontSize:11}} onClick={()=>handleReject(n.rowB)}>B 반려</button>
                      </div>
                    )}
                  </div>
                  <div style={{fontSize:10,color:"#aaa",flexShrink:0}}>
                    {n.unread&&!n.resolved&&<span style={{...bd("red"),padding:"1px 5px",fontSize:9,display:"block",marginBottom:3}}>new</span>}
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
