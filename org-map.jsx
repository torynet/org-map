import { useState, useEffect, useRef } from "react";

const PAL = [
  { bg:'#FDE8E8', tx:'#5C0C0C', ac:'#D83535', dbg:'#5C0C0C', dtx:'#F5B8B8', dac:'#E86868' }, // red
  { bg:'#FDE9E4', tx:'#5C1E08', ac:'#E05022', dbg:'#5C1E08', dtx:'#F5C4AE', dac:'#EC7858' }, // vermillion
  { bg:'#FEF2E6', tx:'#5C3400', ac:'#F09010', dbg:'#5C3400', dtx:'#FAD0A0', dac:'#F8B040' }, // orange
  { bg:'#FEF8E4', tx:'#524200', ac:'#D4B008', dbg:'#524200', dtx:'#F5E090', dac:'#E0CC35' }, // gold
  { bg:'#F2F9E5', tx:'#284400', ac:'#78B815', dbg:'#284400', dtx:'#C8E890', dac:'#96D038' }, // chartreuse
  { bg:'#E6F8E8', tx:'#083A10', ac:'#22A830', dbg:'#083A10', dtx:'#9ADCA4', dac:'#48C858' }, // green
  { bg:'#E6F7F1', tx:'#083E28', ac:'#10A860', dbg:'#083E28', dtx:'#98DCC4', dac:'#30C880' }, // emerald
  { bg:'#E5F7F6', tx:'#063A38', ac:'#0EA8A0', dbg:'#063A38', dtx:'#96D8D5', dac:'#28C8C0' }, // teal
  { bg:'#E5F5FB', tx:'#063038', ac:'#0898C0', dbg:'#063038', dtx:'#96D0E8', dac:'#28B4DC' }, // cyan
  { bg:'#E6F1FC', tx:'#082A50', ac:'#1878D0', dbg:'#082A50', dtx:'#98C8F4', dac:'#40A0EC' }, // sky
  { bg:'#EAF0FE', tx:'#0C1858', ac:'#2850E0', dbg:'#0C1858', dtx:'#A8B8F8', dac:'#5878F0' }, // blue
  { bg:'#EEEBFE', tx:'#180C58', ac:'#5838E0', dbg:'#180C58', dtx:'#C0B5F8', dac:'#8068EC' }, // indigo
  { bg:'#F5E9FE', tx:'#340858', ac:'#9030D8', dbg:'#340858', dtx:'#D8A8F5', dac:'#B058EC' }, // purple
  { bg:'#FAEAFE', tx:'#4A0850', ac:'#C028C0', dbg:'#4A0850', dtx:'#ECA8F5', dac:'#D855D8' }, // violet
  { bg:'#FDEAF8', tx:'#500840', ac:'#CC28A0', dbg:'#500840', dtx:'#F0A8E4', dac:'#E055BC' }, // fuchsia
  { bg:'#FDEAF1', tx:'#580820', ac:'#D82878', dbg:'#580820', dtx:'#F0A8CC', dac:'#E855A0' }, // hot pink
];
const randomPal = () => PAL[Math.floor(Math.random() * PAL.length)];
const unusedPal = (existing) => { const usedAc = new Set(existing.map(x=>x.pal?.ac).filter(Boolean)); const free = PAL.filter(p=>!usedAc.has(p.ac)); const pool = free.length ? free : PAL; return pool[Math.floor(Math.random()*pool.length)]; };
const INIT = { tribes:[], squads:[], chapters:[], guilds:[], people:[], gm:[] };
const uid = () => Math.random().toString(36).slice(2,9);
const dedupById = arr => arr.filter((x,i,a) => a.findIndex(y=>y.id===x.id)===i);
const byName = (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base', numeric: true });

const dedupTribes = (tribes, squads) => {
  const seen={}, idMap={}, kept=[];
  for (const t of tribes) {
    const key=t.name.toLowerCase().trim();
    if (seen[key]!==undefined) idMap[t.id]=seen[key];
    else { seen[key]=t.id; kept.push(t); }
  }
  return { tribes:kept, squads:squads.map(s=>idMap[s.tribeId]?{...s,tribeId:idMap[s.tribeId]}:s) };
};

const migratePerson = ({ squadId, squadIds, chapterId, ...p }) => {
  const sids = p.assignments ? [] : (squadIds?.length ? squadIds : (squadId ? [squadId] : []));
  if (p.assignments?.length) return p;
  const cid = chapterId || null;
  return { ...p, assignments: sids.map(sid => ({ squadId: sid, chapterIds: cid ? [cid] : [] })) };
};

const clean = raw => {
  const base = {
    tribes:   dedupById(raw.tribes   || []),
    squads:   dedupById(raw.squads   || []),
    chapters: dedupById(raw.chapters || []),
    guilds:   dedupById(raw.guilds   || []),
    people:   dedupById((raw.people  || []).map(migratePerson)),
    gm:       [...new Set(raw.gm     || [])],
  };
  const { tribes, squads } = dedupTribes(base.tribes, base.squads);
  const guilds = base.guilds.map(g => g.pal ? g : {...g, pal:randomPal()});
  return { ...base, tribes, squads, guilds };
};

const isEmptyData = d => !d || ['tribes','squads','chapters','guilds','people'].every(k=>d[k].length===0);

const storageGet = (key, shared) => Promise.race([
  window.storage.get(key, shared),
  new Promise((_, rej) => setTimeout(()=>rej(new Error('timeout')), 2500)),
]);

// Solid modal background that works in both light and dark mode
const modalBg  = dk => dk ? '#1e1e1e' : '#ffffff';
const modalBdr = dk => dk ? '#444'    : '#e0e0e0';
const bodyTx   = dk => dk ? '#f0f0f0' : '#111111';
const mutedTx  = dk => dk ? '#aaaaaa' : '#555555';
const inputBg  = dk => dk ? '#2a2a2a' : '#ffffff';
const inputBdr = dk => dk ? '#555'    : '#cccccc';

function EditableLabel({ value, onSave, dk, style={} }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value);
  const ref = useRef();
  useEffect(()=>{ if(editing) ref.current?.select(); },[editing]);
  useEffect(()=>{ setVal(value); },[value]);
  const commit = () => { setEditing(false); const v=val.trim(); if(v&&v!==value) onSave(v); else setVal(value); };
  if (editing) return (
    <input ref={ref} value={val} onChange={e=>setVal(e.target.value)} onBlur={commit}
      onKeyDown={e=>{ if(e.key==='Enter') commit(); if(e.key==='Escape'){setVal(value);setEditing(false);} }}
      style={{fontSize:'inherit',fontFamily:'inherit',fontWeight:'inherit',
        color: dk ? '#f0f0f0' : '#111',
        background: dk ? '#2a2a2a' : '#ffffff',
        border: `1px solid ${dk ? '#666' : '#bbb'}`,
        borderRadius:4, outline:'none', padding:'1px 4px',
        width:Math.max(60,val.length*8)+'px'}}/>
  );
  return <span onDoubleClick={()=>setEditing(true)} title="Double-click to edit" style={{cursor:'text',...style}}>{value}</span>;
}

function RestoreScreen({ onRestore, onStartFresh }) {
  const [text, setText] = useState('');
  const [err, setErr] = useState('');

  const tryParsed = json => {
    const parsed = JSON.parse(json);
    if (!['tribes','squads','chapters','guilds','people','gm'].every(k=>Array.isArray(parsed[k]))) throw new Error();
    onRestore(clean(parsed));
  };

  const tryRestore = () => {
    try { tryParsed(text.trim()); }
    catch { setErr("Could not parse that JSON. Make sure you're pasting a full export from this app."); }
  };

  const onFile = e => {
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try { tryParsed(ev.target.result); }
      catch { setErr("Could not read that file. Make sure it's a previously exported org-map.json."); }
    };
    reader.readAsText(file);
    e.target.value='';
  };

  return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:400,padding:'2rem',fontFamily:'var(--font-sans)'}}>
      <div style={{width:'100%',maxWidth:440}}>
        <p style={{margin:'0 0 6px',fontWeight:500,fontSize:15,color:'var(--color-text-primary)'}}>Restore previous data</p>
        <p style={{margin:'0 0 12px',fontSize:13,color:'var(--color-text-secondary)'}}>Load a file or paste exported JSON, or start fresh.</p>
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10,padding:'8px 10px',
          border:'0.5px solid var(--color-border-tertiary)',borderRadius:'var(--border-radius-md)',
          background:'var(--color-background-secondary)'}}>
          <span style={{fontSize:13,color:'var(--color-text-secondary)',flexShrink:0}}>Load file</span>
          <input type="file" accept=".json" onChange={onFile} style={{fontSize:13,flex:1,minWidth:0}}/>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:8}}>
          <div style={{flex:1,height:'0.5px',background:'var(--color-border-tertiary)'}}/>
          <span style={{fontSize:11,color:'var(--color-text-tertiary)'}}>or paste</span>
          <div style={{flex:1,height:'0.5px',background:'var(--color-border-tertiary)'}}/>
        </div>
        <textarea value={text} onChange={e=>{setText(e.target.value);setErr('');}} placeholder="Paste JSON here…"
          style={{width:'100%',height:130,fontSize:12,fontFamily:'var(--font-mono)',padding:'8px',
            boxSizing:'border-box',resize:'vertical',
            background:'var(--color-background-secondary)',color:'var(--color-text-primary)',
            border:'0.5px solid var(--color-border-tertiary)',borderRadius:'var(--border-radius-md)'}}/>
        {err&&<p style={{margin:'4px 0 0',fontSize:12,color:'var(--color-text-danger)'}}>{err}</p>}
        <div style={{display:'flex',gap:8,marginTop:10,justifyContent:'flex-end'}}>
          <button onClick={onStartFresh} style={{fontSize:13}}>Start fresh</button>
          <button onClick={tryRestore} disabled={!text.trim()} style={{fontSize:13}}>Restore from paste</button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [d, setD] = useState(null);
  const dRef = useRef(null);
  const [phase, setPhase] = useState('loading');
  const [tab, setTab] = useState('matrix');
  const [modal, setModal] = useState(null);
  const [f, setF] = useState({});
  const [editPerson, setEditPerson] = useState(null);
  const [dk, setDk] = useState(false);
  const [importErr, setImportErr] = useState('');
  const [importText, setImportText] = useState('');
  const [exportJson, setExportJson] = useState('');
  const [palettePick, setPalettePick] = useState(null);
  const exportRef = useRef();

  useEffect(()=>{
    const mq=window.matchMedia('(prefers-color-scheme: dark)');
    setDk(mq.matches);
    mq.addEventListener('change', e=>setDk(e.matches));
  },[]);

  useEffect(()=>{
    let cancelled=false;
    (async()=>{
      try {
        const r=await storageGet('om-v2',true);
        if(cancelled)return;
        const loaded=r?clean(JSON.parse(r.value)):null;
        if(loaded&&!isEmptyData(loaded)){ setD(loaded); dRef.current=loaded; setPhase('ready'); }
        else setPhase('restore');
      } catch { if(!cancelled) setPhase('restore'); }
    })();
    return ()=>{ cancelled=true; };
  },[]);

  useEffect(()=>{ dRef.current=d; },[d]);

  const save = nd => {
    const c=clean(nd); setD(c); dRef.current=c;
    try { window.storage.set('om-v2',JSON.stringify(c),true); } catch {}
  };

  const handleRestore = data => { save(data); setPhase('ready'); };
  const handleStartFresh = () => { save(INIT); setPhase('ready'); };

  const ff = p => setF(prev=>({...prev,...p}));
  const cl = (pal,k) => !pal?undefined:dk?pal['d'+k]:pal[k];
  const getTribe = id => d?.tribes.find(t=>t.id===id);
  const getSquad  = id => d?.squads.find(s=>s.id===id);
  const rename = (col,id,name) => save({...d,[col]:d[col].map(x=>x.id===id?{...x,name}:x)});
  const setPal = (col,id,pal) => save({...d,[col]:d[col].map(x=>x.id===id?{...x,pal}:x)});

  const adds = {
    tribe:   ()=>{ if(!f.name?.trim())return; save({...d,tribes:[...d.tribes,{id:uid(),name:f.name.trim(),pal:unusedPal(d.tribes)}]}); setModal(null); },
    squad:   ()=>{ if(!f.name?.trim()||!f.tribeId)return; save({...d,squads:[...d.squads,{id:uid(),name:f.name.trim(),tribeId:f.tribeId}]}); setModal(null); },
    chapter: ()=>{ if(!f.name?.trim())return; save({...d,chapters:[...d.chapters,{id:uid(),name:f.name.trim()}]}); setModal(null); },
    guild:   ()=>{ if(!f.name?.trim())return; save({...d,guilds:[...d.guilds,{id:uid(),name:f.name.trim(),pal:unusedPal(d.guilds)}]}); setModal(null); },
    person:  ()=>{ if(!f.name?.trim()||!f.assignments?.some(a=>a.chapterIds.length>0))return; save({...d,people:[...d.people,{id:uid(),name:f.name.trim(),assignments:f.assignments.filter(a=>a.chapterIds.length>0)}]}); setModal(null); },
  };

  const dels = {
    tribe: id=>{
      const sids=d.squads.filter(s=>s.tribeId===id).map(s=>s.id);
      const np=d.people.map(p=>({...p,assignments:p.assignments.filter(a=>!sids.includes(a.squadId))}));
      const rm=np.filter(p=>p.assignments.length===0).map(p=>p.id);
      save({...d,tribes:d.tribes.filter(t=>t.id!==id),squads:d.squads.filter(s=>s.tribeId!==id),
        people:np.filter(p=>p.assignments.length>0),gm:d.gm.filter(m=>!rm.some(pid=>m.startsWith(pid+'|')))});
    },
    squad: id=>{
      const np=d.people.map(p=>({...p,assignments:p.assignments.filter(a=>a.squadId!==id)}));
      const rm=np.filter(p=>p.assignments.length===0).map(p=>p.id);
      save({...d,squads:d.squads.filter(s=>s.id!==id),people:np.filter(p=>p.assignments.length>0),
        gm:d.gm.filter(m=>!rm.some(pid=>m.startsWith(pid+'|')))});
    },
    chapter: id=>{
      const np=d.people.map(p=>({...p,assignments:p.assignments.map(a=>({...a,chapterIds:a.chapterIds.filter(c=>c!==id)})).filter(a=>a.chapterIds.length>0)}));
      const rm=np.filter(p=>p.assignments.length===0).map(p=>p.id);
      save({...d,chapters:d.chapters.filter(c=>c.id!==id),people:np.filter(p=>p.assignments.length>0),
        gm:d.gm.filter(m=>!rm.some(pid=>m.startsWith(pid+'|')))});
    },
    guild:  id=>save({...d,guilds:d.guilds.filter(g=>g.id!==id),gm:d.gm.filter(m=>!m.endsWith('|'+id))}),
    person: id=>save({...d,people:d.people.filter(p=>p.id!==id),gm:d.gm.filter(m=>!m.startsWith(id+'|'))}),
  };

  const removeFromSquad = (personId, squadId) => {
    const p=d.people.find(x=>x.id===personId); if(!p)return;
    const newAssign=p.assignments.filter(a=>a.squadId!==squadId);
    if(newAssign.length===0) dels.person(personId);
    else save({...d,people:d.people.map(x=>x.id===personId?{...x,assignments:newAssign}:x)});
  };

  const openEditPerson = p => setEditPerson({id:p.id,name:p.name,assignments:p.assignments.map(a=>({squadId:a.squadId,chapterIds:[...a.chapterIds]}))});
  const saveEditPerson = () => {
    const valid=editPerson.assignments.filter(a=>a.chapterIds.length>0);
    if(!editPerson.name.trim()||!valid.length)return;
    save({...d,people:d.people.map(p=>p.id===editPerson.id?{...p,name:editPerson.name.trim(),assignments:valid}:p)});
    setEditPerson(null);
  };
  const toggleEditSquad = sid => setEditPerson(prev=>{
    const has=prev.assignments.some(a=>a.squadId===sid);
    return {...prev,assignments:has?prev.assignments.filter(a=>a.squadId!==sid):[...prev.assignments,{squadId:sid,chapterIds:[]}]};
  });
  const toggleEditChapter = (sid,cid) => setEditPerson(prev=>({
    ...prev,assignments:prev.assignments.map(a=>a.squadId===sid?{...a,chapterIds:a.chapterIds.includes(cid)?a.chapterIds.filter(x=>x!==cid):[...a.chapterIds,cid]}:a)
  }));
  const toggleGM=(pid,gid)=>{ const k=`${pid}|${gid}`; save({...d,gm:d.gm.includes(k)?d.gm.filter(m=>m!==k):[...d.gm,k]}); };

  const doExport=()=>{
    const json=JSON.stringify(dRef.current,null,2);
    setExportJson(json); setModal('export');
    setTimeout(()=>exportRef.current?.select(),50);
  };
  const doDownload=()=>{
    const json=JSON.stringify(dRef.current,null,2);
    const blob=new Blob([json],{type:'application/json'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; a.download='org-map.json'; a.click();
    URL.revokeObjectURL(url);
  };

  const tryImportJson=json=>{
    const parsed=JSON.parse(json);
    if(!['tribes','squads','chapters','guilds','people','gm'].every(k=>Array.isArray(parsed[k])))throw new Error();
    save(clean(parsed)); setImportErr(''); setImportText(''); setModal(null);
  };
  const doImport=e=>{
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try { tryImportJson(ev.target.result); }
      catch { setImportErr("Invalid file — make sure you're importing a previously exported org-map.json."); }
    };
    reader.readAsText(file);
    e.target.value='';
  };
  const doImportPaste=()=>{
    try { tryImportJson(importText.trim()); }
    catch { setImportErr("Could not parse that JSON. Make sure you're pasting a full export from this app."); }
  };

  if(phase==='loading') return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:300,fontFamily:'var(--font-sans)'}}>
      <p style={{color:'var(--color-text-secondary)',fontSize:13}}>Loading…</p>
    </div>
  );
  if(phase==='restore') return <RestoreScreen onRestore={handleRestore} onStartFresh={handleStartFresh}/>;

  const mBg  = modalBg(dk);
  const mBdr = modalBdr(dk);
  const mTx  = bodyTx(dk);
  const mMut = mutedTx(dk);
  const iBg  = inputBg(dk);
  const iBdr = inputBdr(dk);
  const iSt  = {width:'100%',boxSizing:'border-box',fontFamily:'inherit',fontSize:13,
    background:iBg,color:mTx,border:`1px solid ${iBdr}`,borderRadius:6,padding:'6px 8px',outline:'none'};
  const btnSt = {fontSize:13,padding:'5px 14px',borderRadius:6,cursor:'pointer',fontFamily:'inherit',
    background:iBg,color:mTx,border:`1px solid ${iBdr}`};
  const btnPrimary = {...btnSt,background:dk?'#555':'#222',color:'#fff',border:'none'};

  const tribesWithSquads=d.tribes.map(t=>({...t,squads:d.squads.filter(s=>s.tribeId===t.id)})).filter(t=>t.squads.length>0);
  const flatSquads=tribesWithSquads.flatMap(t=>t.squads.map((s,i)=>({s,t,first:i===0,last:i===t.squads.length-1})));
  const cellPeople=(chId,sqId)=>d.people.filter(p=>p.assignments.some(a=>a.squadId===sqId&&a.chapterIds.includes(chId)));
  const personGuilds=pid=>d.guilds.filter(g=>d.gm.includes(`${pid}|${g.id}`));
  const personPal=p=>{ const sq=getSquad(p.assignments?.[0]?.squadId); return sq?getTribe(sq.tribeId)?.pal:null; };
  const open=(key,defs={})=>{ setF({name:'',...defs}); setImportErr(''); setImportText(''); setModal(key); };
  const toggleSquadSel=sid=>{ const cur=f.assignments||[]; ff({assignments:cur.some(a=>a.squadId===sid)?cur.filter(a=>a.squadId!==sid):[...cur,{squadId:sid,chapterIds:[]}]}); };
  const toggleChapterSel=(sid,cid)=>{ ff({assignments:(f.assignments||[]).map(a=>a.squadId===sid?{...a,chapterIds:a.chapterIds.includes(cid)?a.chapterIds.filter(x=>x!==cid):[...a.chapterIds,cid]}:a)}); };

  const rowDivider = dk ? '1px solid rgba(255,255,255,0.25)' : '1px solid rgba(0,0,0,0.2)';
  const cellStyle=(t,first,last,extra={},rowIdx=0)=>({
    borderBottom: rowDivider,
    borderLeft:  first?`2px solid ${cl(t.pal,'ac')}`:`0.5px solid ${cl(t.pal,'ac')}44`,
    borderRight: last ?`2px solid ${cl(t.pal,'ac')}`:'none',
    backgroundColor: dk?`${t.pal.dbg}44`:`${t.pal.bg}cc`,
    backgroundImage: rowIdx%2===1
      ? 'linear-gradient(rgba(0,0,0,0.09),rgba(0,0,0,0.09))'
      : 'none',
    ...extra,
  });;

  const SquadChecklist = ({ selectedIds, onToggle }) => (
    <div style={{border:`1px solid ${iBdr}`,borderRadius:6,maxHeight:160,overflowY:'auto',padding:'4px',background:iBg}}>
      {[...d.tribes].sort(byName).map(t=>{
        const ts=[...d.squads].filter(s=>s.tribeId===t.id).sort(byName);
        if(!ts.length)return null;
        return (
          <div key={t.id}>
            <p style={{margin:'4px 6px 2px',fontSize:10,fontWeight:500,color:mMut,letterSpacing:'0.05em',textTransform:'uppercase'}}>{t.name}</p>
            {ts.map(s=>(
              <label key={s.id} style={{display:'flex',alignItems:'center',gap:6,padding:'4px 6px',cursor:'pointer',borderRadius:4,userSelect:'none'}}>
                <input type="checkbox" checked={selectedIds.includes(s.id)} onChange={()=>onToggle(s.id)} style={{margin:0,cursor:'pointer'}}/>
                <span style={{fontSize:13,color:mTx}}>{s.name}</span>
              </label>
            ))}
          </div>
        );
      })}
      {d.squads.length===0&&<p style={{margin:'8px',fontSize:12,color:mMut}}>No squads yet.</p>}
    </div>
  );

  const AssignmentEditor = ({ assignments, onToggleSquad, onToggleChapter }) => (
    <div style={{border:`1px solid ${iBdr}`,borderRadius:6,maxHeight:280,overflowY:'auto',padding:'4px',background:iBg}}>
      {[...d.tribes].sort(byName).map(t=>{
        const ts=[...d.squads].filter(s=>s.tribeId===t.id).sort(byName);
        if(!ts.length)return null;
        return (
          <div key={t.id}>
            <p style={{margin:'4px 6px 2px',fontSize:10,fontWeight:500,color:mMut,letterSpacing:'0.05em',textTransform:'uppercase'}}>{t.name}</p>
            {ts.map(s=>{
              const assign=assignments.find(a=>a.squadId===s.id);
              const selected=!!assign;
              return (
                <div key={s.id}>
                  <label style={{display:'flex',alignItems:'center',gap:6,padding:'4px 6px',cursor:'pointer',borderRadius:4,userSelect:'none'}}>
                    <input type="checkbox" checked={selected} onChange={()=>onToggleSquad(s.id)} style={{margin:0,cursor:'pointer'}}/>
                    <span style={{fontSize:13,color:mTx,fontWeight:selected?500:400}}>{s.name}</span>
                  </label>
                  {selected&&d.chapters.length>0&&(
                    <div style={{marginLeft:24,marginBottom:4}}>
                      {[...d.chapters].sort(byName).map(c=>(
                        <label key={c.id} style={{display:'flex',alignItems:'center',gap:6,padding:'2px 6px',cursor:'pointer',borderRadius:4,userSelect:'none'}}>
                          <input type="checkbox" checked={assign.chapterIds.includes(c.id)} onChange={()=>onToggleChapter(s.id,c.id)} style={{margin:0,cursor:'pointer'}}/>
                          <span style={{fontSize:12,color:mMut}}>{c.name}</span>
                        </label>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        );
      })}
      {d.squads.length===0&&<p style={{margin:'8px',fontSize:12,color:mMut}}>No squads yet.</p>}
    </div>
  );

  // Modal shell with explicit solid colors
  const Modal = ({title, onClose, children, extra=null}) => (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:'absolute',inset:0,background:'rgba(0,0,0,0.5)',
        display:'flex',alignItems:'center',justifyContent:'center',zIndex:20}}>
      <div style={{background:mBg,border:`1px solid ${mBdr}`,borderRadius:12,
        padding:'1.25rem',width:340,display:'flex',flexDirection:'column',gap:10,
        boxShadow:'0 8px 32px rgba(0,0,0,0.4)'}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <p style={{margin:0,fontWeight:500,fontSize:14,color:mTx}}>{title}</p>
          {extra}
        </div>
        {children}
      </div>
    </div>
  );

  const PalettePickerModal = ({item, col, onClose}) => {
    if(!item) return null;
    const getUnused = () => {
      const usedAc = new Set(d[col].filter(x=>x.id!==item.id).map(x=>x.pal?.ac).filter(Boolean));
      const free = PAL.filter(p=>!usedAc.has(p.ac));
      return free.length ? free : PAL;
    };
    return (
      <Modal title="Choose color" onClose={onClose}>
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:6}}>
          {PAL.map((p)=>(
            <button key={p.ac} onClick={()=>{setPal(col,item.id,p);onClose();}}
              style={{width:'100%',aspectRatio:'1',borderRadius:8,border:item.pal===p?`3px solid ${p.ac}`:`2px solid ${p.ac}66`,
                background:p.ac,cursor:'pointer',transition:'all 0.15s',transform:item.pal===p?'scale(1.1)':'scale(1)',
                boxShadow:item.pal===p?`0 2px 8px ${p.ac}88`:'none'}}/>
          ))}
        </div>
        <button onClick={()=>{const pool=getUnused();setPal(col,item.id,pool[Math.floor(Math.random()*pool.length)]);onClose();}}
          style={{...btnPrimary,marginTop:4}}>Random (unused)</button>
      </Modal>
    );
  };

  return (
    <>
      <style>{`
        .delbtn{opacity:0;font-size:11px;padding:0 3px;border:none;background:none;color:var(--color-text-tertiary);cursor:pointer;line-height:1;flex-shrink:0}
        .hovdel:hover .delbtn,.pcard:hover .delbtn{opacity:1}
      `}</style>

      <div style={{padding:'1rem',position:'relative',minHeight:500,fontFamily:'var(--font-sans)'}}>

        <div style={{display:'flex',gap:8,marginBottom:'1rem',flexWrap:'wrap',alignItems:'center'}}>
          <div style={{display:'flex',border:'0.5px solid var(--color-border-tertiary)',borderRadius:'var(--border-radius-md)',overflow:'hidden',flexShrink:0}}>
            {['matrix','guilds'].map(t=>(
              <button key={t} onClick={()=>setTab(t)} style={{fontSize:13,padding:'5px 14px',border:'none',cursor:'pointer',fontFamily:'inherit',background:tab===t?'var(--color-background-secondary)':'transparent',color:tab===t?'var(--color-text-primary)':'var(--color-text-secondary)'}}>
                {t.charAt(0).toUpperCase()+t.slice(1)}
              </button>
            ))}
          </div>
          <div style={{flex:1}}/>
          {[['+ Tribe','tribe',{}],['+ Squad','squad',{tribeId:[...d.tribes].sort(byName)[0]?.id||''}],['+ Chapter','chapter',{}],['+ Guild','guild',{}],['+ Person','person',{assignments:[]}]].map(([label,key,defs])=>(
            <button key={key} onClick={()=>open(key,defs)} style={{fontSize:12}}>{label}</button>
          ))}
          <div style={{width:'0.5px',background:'var(--color-border-tertiary)',alignSelf:'stretch'}}/>
          <button onClick={doExport} style={{fontSize:12}}>Export</button>
          <button onClick={()=>open('import')} style={{fontSize:12}}>Import</button>
        </div>

        {tab==='matrix'&&(
          flatSquads.length===0||d.chapters.length===0
            ?<Empty>Add tribes, squads, and chapters to see the matrix.</Empty>
            :<div style={{overflowX:'auto'}}>
              <table style={{borderCollapse:'collapse',fontSize:13}}>
                <thead>
                  <tr>
                    <td style={{minWidth:110,borderBottom:'0.5px solid var(--color-border-secondary)'}}/>
                    {tribesWithSquads.map(t=>(
                      <th key={t.id} colSpan={t.squads.length}
                        style={{background:cl(t.pal,'bg'),color:cl(t.pal,'tx'),padding:'6px 10px',textAlign:'center',
                          fontSize:11,fontWeight:500,letterSpacing:'0.04em',whiteSpace:'nowrap',
                          borderLeft:`2px solid ${cl(t.pal,'ac')}`,borderRight:`2px solid ${cl(t.pal,'ac')}`,borderTop:`2px solid ${cl(t.pal,'ac')}`}}>
                        <div className="hovdel" style={{display:'flex',alignItems:'center',justifyContent:'center',gap:4}}>
                          <EditableLabel value={t.name} onSave={v=>rename('tribes',t.id,v)} dk={dk} style={{color:cl(t.pal,'tx')}}/>
                          <button onClick={()=>setPalettePick({col:'tribes',item:t})} style={{width:16,height:16,borderRadius:3,border:'none',background:t.pal.ac,cursor:'pointer',padding:0,opacity:0.7}}title="Change color"/>
                          <button className="delbtn" onClick={()=>dels.tribe(t.id)}>×</button>
                        </div>
                      </th>
                    ))}
                  </tr>
                  <tr>
                    <th style={{padding:'3px 8px',textAlign:'left',color:'var(--color-text-secondary)',fontSize:11,fontWeight:400,borderBottom:'0.5px solid var(--color-border-secondary)'}}/>
                    {flatSquads.map(({s,t,first,last})=>(
                      <th key={s.id} style={{...cellStyle(t,first,last,{padding:'5px 8px',textAlign:'center',borderBottom:'0.5px solid var(--color-border-secondary)'}),minWidth:100}}>
                        <div className="hovdel" style={{display:'flex',flexDirection:'column',alignItems:'center',gap:1}}>
                          <span style={{fontSize:12,fontWeight:400,color:'var(--color-text-primary)',maxWidth:96,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',display:'block'}}>
                            <EditableLabel value={s.name} onSave={v=>rename('squads',s.id,v)} dk={dk}/>
                          </span>
                          <div style={{display:'flex',gap:2,alignItems:'center'}}>
                            <button className="delbtn" onClick={()=>dels.squad(s.id)}>×</button>
                          </div>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {d.chapters.map((ch,i)=>(
                    <tr key={ch.id}>
                      <td style={{padding:'6px 8px',whiteSpace:'nowrap',background:i%2===1?'var(--color-background-secondary)':'transparent',borderBottom:rowDivider,borderRight:'0.5px solid var(--color-border-tertiary)'}}>
                        <div className="hovdel" style={{display:'flex',alignItems:'center',gap:4}}>
                          <span style={{fontSize:12,color:'var(--color-text-secondary)',fontWeight:500}}>
                            <EditableLabel value={ch.name} onSave={v=>rename('chapters',ch.id,v)} dk={dk}/>
                          </span>
                          <button className="delbtn" onClick={()=>dels.chapter(ch.id)}>×</button>
                        </div>
                      </td>
                      {flatSquads.map(({s,t,first,last})=>(
                        <td key={s.id} style={{...cellStyle(t,first,last,{padding:'4px 5px',verticalAlign:'top'},i)}}> 
                          <div style={{display:'flex',flexDirection:'column',gap:3}}>
                            {cellPeople(ch.id,s.id).map(p=>(
                              <PersonCard key={p.id} person={p} guilds={personGuilds(p.id)}
                                onEdit={()=>openEditPerson(p)} onRemove={()=>removeFromSquad(p.id,s.id)} dk={dk}/>
                            ))}
                          </div>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
        )}

        {tab==='guilds'&&(
          d.guilds.length===0
            ?<Empty>No guilds yet. Click '+ Guild' to create one.</Empty>
            :<div style={{display:'flex',flexDirection:'column',gap:10}}>
              {d.guilds.map(guild=>{
                const members=d.people.filter(p=>d.gm.includes(`${p.id}|${guild.id}`));
                const others=d.people.filter(p=>!d.gm.includes(`${p.id}|${guild.id}`));
                const hBg=cl(guild.pal,'bg'),hTx=cl(guild.pal,'tx'),hAc=cl(guild.pal,'ac');
                return (
                  <div key={guild.id} style={{border:`1.5px solid ${hAc}`,borderRadius:'var(--border-radius-lg)',overflow:'hidden'}}>
                    <div style={{background:hBg,padding:'6px 12px',display:'flex',alignItems:'center',gap:6}}>
                      <span style={{fontWeight:500,fontSize:13,color:hTx,flex:1}}>
                        <EditableLabel value={guild.name} onSave={v=>rename('guilds',guild.id,v)} dk={dk} style={{color:hTx}}/>
                      </span>
                      <span style={{fontSize:11,color:hTx,opacity:0.7}}>({members.length} member{members.length!==1?'s':''})</span>
                      <button onClick={()=>setPalettePick({col:'guilds',item:guild})} style={{width:16,height:16,borderRadius:3,border:'1px solid',borderColor:hAc,background:hBg,cursor:'pointer',padding:0,opacity:0.7}}title="Change color"/>
                      <button className="delbtn" onClick={()=>dels.guild(guild.id)} style={{opacity:0.5,color:hTx}}>×</button>
                    </div>
                    <div style={{padding:'8px 12px',display:'flex',flexWrap:'wrap',gap:5,alignItems:'center',background:'var(--color-background-primary)'}}>
                      {members.map(p=>{
                        const sq=getSquad(p.assignments?.[0]?.squadId),tr=sq?getTribe(sq.tribeId):null;
                        return <Chip key={p.id} name={p.name} onEdit={()=>openEditPerson(p)} sub={tr?tr.name:null} pal={personPal(p)} onDel={()=>toggleGM(p.id,guild.id)} dk={dk}/>;
                      })}
                      {others.length>0&&(
                        <select onChange={e=>{if(e.target.value){toggleGM(e.target.value,guild.id);e.target.value=''}}}
                          style={{fontSize:12,padding:'3px 6px',background:'transparent',fontFamily:'inherit',
                            color:'var(--color-text-tertiary)',cursor:'pointer',
                            border:`0.5px dashed ${hAc}`,borderRadius:'var(--border-radius-md)'}}>
                          <option value="">+ add member…</option>
                          {[...others].sort(byName).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                        </select>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
        )}

        {editPerson&&(
          <Modal title="Edit person" onClose={()=>setEditPerson(null)}
            extra={
              <button onClick={()=>{ if(window.confirm('Delete this person entirely?')){ dels.person(editPerson.id); setEditPerson(null); } }}
                style={{fontSize:12,color:'#cc3333',border:'1px solid #cc333366',background:'none',cursor:'pointer',padding:'3px 8px',borderRadius:6}}>
                Delete
              </button>
            }>
            <input placeholder="Name" value={editPerson.name}
              onChange={e=>setEditPerson(p=>({...p,name:e.target.value}))}
              style={iSt}/>
            <p style={{margin:0,fontSize:12,color:mMut}}>Squads & Chapters</p>
            <AssignmentEditor assignments={editPerson.assignments}
              onToggleSquad={toggleEditSquad}
              onToggleChapter={toggleEditChapter}/>
            <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
              <button onClick={()=>setEditPerson(null)} style={btnSt}>Cancel</button>
              <button onClick={saveEditPerson} style={btnPrimary}>Save</button>
            </div>
          </Modal>
        )}

        {modal&&(
          <Modal
            title={{tribe:'Add tribe',squad:'Add squad',chapter:'Add chapter',guild:'Add guild',person:'Add person',import:'Import data',export:'Export data'}[modal]}
            onClose={()=>setModal(null)}>
            {modal==='export'&&(<>
              <p style={{margin:0,fontSize:12,color:mMut}}>Download the file or copy from the box below.</p>
              <textarea ref={exportRef} readOnly value={exportJson} onClick={()=>exportRef.current?.select()}
                style={{width:'100%',height:200,fontSize:11,fontFamily:'var(--font-mono)',resize:'vertical',
                  background:iBg,color:mTx,border:`1px solid ${iBdr}`,borderRadius:6,padding:'8px',
                  boxSizing:'border-box'}}/>
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                <button onClick={()=>navigator.clipboard.writeText(exportJson)} style={btnSt}>Copy</button>
                <button onClick={doDownload} style={btnSt}>Download</button>
                <button onClick={()=>setModal(null)} style={btnSt}>Close</button>
              </div>
            </>)}
            {modal==='import'&&(<>
              <p style={{margin:0,fontSize:13,color:mMut}}>Load a file or paste exported JSON. Replaces current data.</p>
              {importErr&&<p style={{margin:0,fontSize:12,color:'#cc3333'}}>{importErr}</p>}
              <input type="file" accept=".json" onChange={doImport} style={{fontSize:13,color:mTx}}/>
              <p style={{margin:0,fontSize:12,color:mMut,textAlign:'center'}}>— or paste JSON —</p>
              <textarea value={importText} onChange={e=>setImportText(e.target.value)} placeholder="Paste exported JSON here…"
                style={{width:'100%',height:120,fontSize:11,fontFamily:'var(--font-mono)',resize:'vertical',
                  background:iBg,color:mTx,border:`1px solid ${iBdr}`,borderRadius:6,padding:'8px',
                  boxSizing:'border-box'}}/>
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                <button onClick={()=>setModal(null)} style={btnSt}>Cancel</button>
                <button onClick={doImportPaste} disabled={!importText.trim()} style={btnPrimary}>Import</button>
              </div>
            </>)}
            {!['export','import'].includes(modal)&&(<>
              <input autoFocus placeholder="Name" value={f.name||''}
                onChange={e=>ff({name:e.target.value})}
                onKeyDown={e=>e.key==='Enter'&&modal!=='person'&&adds[modal]?.()}
                style={iSt}/>
              {modal==='squad'&&(
                <select value={f.tribeId||''} onChange={e=>ff({tribeId:e.target.value})} style={iSt}>
                  <option value="">Select tribe…</option>
                  {[...d.tribes].sort(byName).map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              )}
              {modal==='person'&&(<>
                <p style={{margin:0,fontSize:12,color:mMut}}>Squads & Chapters</p>
                <AssignmentEditor assignments={f.assignments||[]}
                  onToggleSquad={toggleSquadSel}
                  onToggleChapter={toggleChapterSel}/>
              </>)}
              <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
                <button onClick={()=>setModal(null)} style={btnSt}>Cancel</button>
                <button onClick={adds[modal]} style={btnPrimary}>Add</button>
              </div>
            </>)}
          </Modal>
        )}

        {palettePick&&(
          <PalettePickerModal item={palettePick.item} col={palettePick.col} onClose={()=>setPalettePick(null)}/>
        )}
      </div>
    </>
  );
}

function Empty({children}){
  return <div style={{textAlign:'center',padding:'3rem 1rem',color:'var(--color-text-secondary)',fontSize:13}}>{children}</div>;
}

function PersonCard({ person, guilds, onEdit, onRemove, dk }) {
  return (
    <div className="pcard" style={{display:'inline-flex',flexDirection:'column',gap:2,padding:'3px 6px',
      background:'var(--color-background-primary)',border:'0.5px solid var(--color-border-tertiary)',
      borderRadius:6,maxWidth:150,minWidth:60}}>
      <div style={{display:'flex',alignItems:'center',gap:3}}>
        <span onClick={onEdit} title="Click to edit"
          style={{fontSize:12,color:'var(--color-text-primary)',lineHeight:1.4,flex:1,minWidth:0,
            overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',cursor:'pointer'}}>
          {person.name}
        </span>
        <button className="delbtn" onClick={onRemove} title="Remove from this squad" style={{marginLeft:2}}>×</button>
      </div>
      {guilds.length>0&&(
        <div style={{display:'flex',flexWrap:'wrap',gap:2}}>
          {guilds.map(g=>{
            const gbg=dk?g.pal.dbg:g.pal.bg, gtx=dk?g.pal.dtx:g.pal.tx, gbo=dk?g.pal.dac:g.pal.ac;
            return <span key={g.id} style={{fontSize:9,padding:'1px 5px',borderRadius:20,background:gbg,
              color:gtx,border:`0.5px solid ${gbo}`,whiteSpace:'nowrap',lineHeight:1.6,
              maxWidth:90,overflow:'hidden',textOverflow:'ellipsis',display:'block'}}>{g.name}</span>;
          })}
        </div>
      )}
    </div>
  );
}

function Chip({name,onEdit,sub,pal,onDel,dk}){
  const bg=pal?(dk?pal.dbg:pal.bg):'var(--color-background-secondary)';
  const tx=pal?(dk?pal.dtx:pal.tx):'var(--color-text-primary)';
  const bo=pal?(dk?pal.dac:pal.ac):'var(--color-border-tertiary)';
  const st=pal?(dk?pal.dac:pal.ac):'var(--color-text-tertiary)';
  return (
    <div className="pcard" style={{display:'inline-flex',alignItems:'center',gap:4,padding:'2px 6px',
      background:bg,border:`0.5px solid ${bo}`,borderRadius:6,maxWidth:160,flexShrink:0}}>
      <div style={{display:'flex',flexDirection:'column',minWidth:0,overflow:'hidden',flex:1}}>
        <span onClick={onEdit} title="Click to edit" style={{fontSize:12,color:tx,lineHeight:1.4,cursor:'pointer'}}>{name}</span>
        {sub&&<span style={{fontSize:10,color:st,lineHeight:1.2,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}}>{sub}</span>}
      </div>
      <button className="delbtn" onClick={onDel}
        style={{opacity:0.4,fontSize:11,padding:'0 1px',border:'none',background:'none',
          cursor:'pointer',color:tx,flexShrink:0,lineHeight:1}}
        onMouseEnter={e=>e.currentTarget.style.opacity=1}
        onMouseLeave={e=>e.currentTarget.style.opacity=0.4}>×</button>
    </div>
  );
}
