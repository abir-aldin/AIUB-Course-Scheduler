let rows=[],groups={},selected=new Set(),generated=[],fixedSelections=[];
const $=id=>document.getElementById(id);

function norm(v){return String(v??'').trim().toLowerCase().replace(/[^a-z0-9]+/g,'')}
function findKey(obj,names){
  const keys=Object.keys(obj), want=names.map(norm);
  return keys.find(k=>want.includes(norm(k))) ||
         keys.find(k=>want.some(w=>norm(k).includes(w))) || null;
}
function num(v){
  if(typeof v==='number') return v;
  const m=String(v??'').replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);
  return m ? Number(m[0]) : 0;
}
function esc(v){
  return String(v??'').replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
}
function parseCourseSection(v){
  const s=String(v??'').trim();
  let m=s.match(/^(.*?)\s*[\[\(]\s*(?:section\s*)?([A-Za-z0-9_-]+)\s*[\]\)]\s*$/i);
  if(m) return {course:m[1].trim(),section:m[2].trim()};
  m=s.match(/^(.*?)\s*[-–—]\s*(?:section\s*)?([A-Za-z0-9_-]+)\s*$/i);
  if(m) return {course:m[1].trim(),section:m[2].trim()};
  return {course:s,section:''};
}
function findExactKey(obj,names){
  const keys=Object.keys(obj);
  const wanted=names.map(norm);
  return keys.find(k=>wanted.includes(norm(k))) || null;
}

function mapRow(r){
  const ck=findKey(r,['coursecode','code','course id','courseid','course']);
  const nk=findKey(r,['coursename','name','course title','title','course name']);
  const sk=findKey(r,['section','sec','section no','sectionno']);
  const cnt=findKey(r,['count','student count','students','enrolled','enrollment']);
  const dk=findKey(r,['day','days','class day','class days']);
  const tk=findKey(r,['time','timeslot','time slot','class time','schedule']);
  const startKey=findExactKey(r,['start time','starttime','class start','begin time']);
  const endKey=findExactKey(r,['end time','endtime','class end','finish time']);
  const fk=findKey(r,['faculty','teacher','instructor']);
  const rk=findKey(r,['room','room no','room number']);
  const stk=findKey(r,['status','section status','class status','availability']);

  const idk=findExactKey(r,['class id','classid']);
  const typek=findExactKey(r,['type','course type','class type']);

  const codeRaw=String(ck?r[ck]:'').trim();
  const nameRaw=String(nk?r[nk]:'').trim();
  const explicit=String(sk?r[sk]:'').trim();

  const pName=parseCourseSection(nameRaw || codeRaw);
  const pCode=parseCourseSection(codeRaw);

  const course=pName.course || pCode.course;
  const code=pCode.course || codeRaw || course;
  const section=explicit || pName.section || pCode.section;

  const rawTime=String(tk?r[tk]:'').trim();
  const start=String(startKey?r[startKey]:'').trim();
  const end=String(endKey?r[endKey]:'').trim();

  // Preserve separate start/end columns when the Excel has them.
  // Otherwise use the combined Time/Timeslot cell.
  const time=(start && end) ? `${start} - ${end}` : rawTime;

  return {
    classId:String(idk?r[idk]:'').trim(),
    type:String(typek?r[typek]:'').trim(),
    code,course,section,
    count:num(cnt?r[cnt]:''),
    days:String(dk?r[dk]:'').trim(),
    time,
    faculty:String(fk?r[fk]:'').trim(),
    room:String(rk?r[rk]:'').trim(),
    status:(String(stk?r[stk]:'').trim() || 'Open')
  };
}

// Combine all Excel rows belonging to the same Class ID
function aggregateSectionRows(sectionRows){

  const grouped=new Map();

  sectionRows.forEach((r,index)=>{

    const key=r.classId
      ? `id:${r.classId}`
      : `row:${index}`;

    if(!grouped.has(key)){

      grouped.set(key,{
        ...r,
        meetings:[]
      });

    }

    const option=grouped.get(key);

    option.meetings.push({

      days:r.days,
      time:r.time,
      type:r.type,
      faculty:r.faculty,
      room:r.room

    });

  });

  return [...grouped.values()];
}

async function loadFile(file){
  try{
    if(!file)return;
    if(!/\.(xlsx|xls|csv)$/i.test(file.name)){
      alert('Please select an .xlsx, .xls, or .csv file.');
      return;
    }
    if(typeof XLSX==='undefined'){
      alert('The Excel reader did not load. Please keep internet connected, reload the page, and try again.');
      return;
    }

    const buffer=await file.arrayBuffer();
    const wb=XLSX.read(buffer,{type:'array',cellDates:false});
    if(!wb.SheetNames.length)throw new Error('No worksheet found.');

    const ws=wb.Sheets[wb.SheetNames[0]];
    const data=XLSX.utils.sheet_to_json(ws,{defval:'',raw:false});
    if(!data.length)throw new Error('The first worksheet is empty.');

    rows=data.map(mapRow).filter(r=>r.code||r.course||r.section);
    if(!rows.length)throw new Error('No recognizable course rows found.');

    groups={};
    for(const r of rows){
      const groupKey=norm(r.code||r.course);
      if(groupKey)(groups[groupKey]??=[]).push(r);
    }

    // Convert raw Excel rows into complete section options
for(const [groupKey,sectionRows] of Object.entries(groups)){

  groups[groupKey]=aggregateSectionRows(sectionRows);

}

    selected.clear();
    generated=[];
    $('fname').textContent=`Loaded: ${file.name} • ${rows.length} rows • ${Object.keys(groups).length} courses`;
    $('setup').classList.remove('hidden');
    renderCourses();
    renderFixedRows();
  }catch(err){
    console.error(err);
    alert('Could not read Excel: '+(err.message||err));
  }
}

function isOpen(r){
  // Only exact Open status is eligible; case/whitespace differences are ignored.
  return String(r.status||'Open').trim().toLowerCase()==='open';
}

function renderCourses(){
  const q=$('search').value.trim().toLowerCase();
  const min=Number($('minCount').value||0);
  const max=Number($('maxCount').value||Infinity);
  const box=$('courses');
  box.innerHTML='';

  for(const [groupKey,rs] of Object.entries(groups)){
    const matches=!q ||
      groupKey.includes(norm(q)) ||
      rs.some(r=>r.code.toLowerCase().includes(q)||r.course.toLowerCase().includes(q));

    if(!matches)continue;

    const eligible=rs.filter(r=>isOpen(r) && r.count>=min && r.count<=max);
    const label=document.createElement('label');
    label.className='course';
    label.dataset.group=groupKey;

    label.innerHTML=
      `<input type="checkbox" ${selected.has(groupKey)?'checked':''}>
       <div class="name">
         <div class="code">${esc(rs[0].code||rs[0].course)}</div>
         <div class="title">${esc(rs[0].course)}</div>
       </div>
       <span class="badge">${eligible.length}/${rs.length} eligible</span>`;

    label.querySelector('input').addEventListener('change',e=>{
      if(e.target.checked)selected.add(groupKey);
      else selected.delete(groupKey);
      updateSelected();
    });

    box.appendChild(label);
  }
  updateSelected();
}

function updateSelected(){
  $('selected').textContent=`${selected.size} course${selected.size===1?'':'s'} selected`;
}

function timeValue(v){
  let s=String(v??'').trim().toUpperCase().replace(/\./g,':');
  if(!s)return null;

  // 12-hour format: 8 AM, 8:30 AM, 08:30PM
  let m=s.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/);
  if(m){
    let h=Number(m[1]),mi=Number(m[2]||0);
    if(h<1||h>12||mi>59)return null;
    if(m[3]==='PM'&&h<12)h+=12;
    if(m[3]==='AM'&&h===12)h=0;
    return h*60+mi;
  }

  // 24-hour format: 08:30, 13:45
  m=s.match(/^(\d{1,2})(?::(\d{2}))?$/);
  if(m){
    const h=Number(m[1]),mi=Number(m[2]||0);
    if(h<=23&&mi<=59)return h*60+mi;
  }
  return null;
}

function extractTimeParts(v){
  let s=String(v??'').trim().toUpperCase();
  if(!s)return [];

  s=s.replace(/[–—−]/g,'-')
     .replace(/\bUNTIL\b/gi,'-')
     .replace(/\bTO\b/gi,'-')
     .replace(/\bTILL\b/gi,'-')
     .replace(/\s+/g,' ');

  // Normalize 08.30 to 08:30.
  s=s.replace(/(\d{1,2})\.(\d{2})/g,'$1:$2');

  // Find actual time tokens anywhere in the cell.
  const matches=[...s.matchAll(/\b\d{1,2}(?::\d{2})?\s*(?:AM|PM)?\b/g)]
    .map(m=>m[0].trim());

  return matches.filter(x=>timeValue(x)!==null);
}

function timeRange(v){
  const parts=extractTimeParts(v);
  if(parts.length<2)return null;

  const a=timeValue(parts[0]);
  const b=timeValue(parts[1]);
  if(a===null||b===null)return null;

  // If end is earlier than start, treat it as an overnight interval.
  return [a,b<=a?b+1440:b];
}

function formatTimeValue(v){
  const m=Number(v);
  if(!Number.isFinite(m))return '';
  const h24=Math.floor((m%1440)/60), min=m%60;
  const ap=h24>=12?'PM':'AM';
  const h=h24%12||12;
  return `${String(h).padStart(2,'0')}:${String(min).padStart(2,'0')} ${ap}`;
}

function displayTime(v){
  const r=timeRange(v);
  if(!r)return String(v??'-');
  return `${formatTimeValue(r[0])} – ${formatTimeValue(r[1])}`;
}

function displayScheduleMeetings(r){

  const entries=Array.isArray(r.meetings) && r.meetings.length
    ? r.meetings
    : [r];

  return entries.map(m=>{

    const day=esc(m.days||'-');
    const time=esc(displayTime(m.time)||'-');

    const type=m.type
      ? ` <span class="pill">${esc(m.type)}</span>`
      : '';

    return `
      <div class="schedule-meeting">
        <strong>${day}</strong><br>
        ${time}${type}
      </div>
    `;

  }).join('<hr>');

}


// Build weekly table for a generated schedule
function buildWeeklyTable(schedule){

  const weekDays=[
    {key:'SU',label:'Sunday'},
    {key:'M',label:'Monday'},
    {key:'T',label:'Tuesday'},
    {key:'W',label:'Wednesday'},
    {key:'R',label:'Thursday'},
    {key:'F',label:'Friday'},
    {key:'S',label:'Saturday'}
  ];

  const byDay=Object.fromEntries(
    weekDays.map(day=>[day.key,[]])
  );

  schedule.forEach(course=>{

  const entries=Array.isArray(course.meetings) && course.meetings.length
    ? course.meetings
    : [course];

  entries.forEach(meeting=>{

    normalizeDayTokens(meeting.days).forEach(day=>{

      if(byDay[day]){

        byDay[day].push({
          ...course,
          ...meeting
        });

      }

    });

  });

});

  const body=weekDays.map(day=>{

    const classes=[...byDay[day.key]].sort((a,b)=>{
      const timeA=timeRange(a.time);
      const timeB=timeRange(b.time);

      return (timeA?timeA[0]:Infinity)-
             (timeB?timeB[0]:Infinity);
    });

    const scheduleHtml=classes.length
      ? `<div class="weekly-entries">
          ${classes.map(course=>`
            <div class="weekly-entry">

              <div class="weekly-entry-time">
                ${esc(displayTime(course.time))}
              </div>

              <div class="weekly-entry-course">
                <strong>${esc(course.code||course.course||'Course')}</strong>
                ${course.course&&course.course!==course.code
                  ? `<span>${esc(course.course)}</span>`
                  : ''}
              </div>

              <div class="weekly-entry-details">
                Section: <strong>${esc(course.section||'-')}</strong>
                <span>•</span>
                Section Count: <strong>${esc(course.count)}</strong>
              </div>

            </div>
          `).join('')}
        </div>`
      : '<span class="weekly-no-class">No classes</span>';

    return `
      <tr>
        <th scope="row">${day.label}</th>
        <td>${scheduleHtml}</td>
      </tr>
    `;

  }).join('');

  return `
    <table class="weekly-table">

      <thead>
        <tr>
          <th>Day</th>
          <th>All Time Schedule</th>
        </tr>
      </thead>

      <tbody>
        ${body}
      </tbody>

    </table>
  `;
}


// Open weekly schedule popup
function openWeeklyModal(schedule,scheduleNumber){

  $('weeklyModalTitle').textContent =
    `Schedule ${scheduleNumber} — Weekly Table`;

  $('weeklyTableContainer').innerHTML =
    buildWeeklyTable(schedule);

  $('weeklyModal').classList.remove('hidden');

  document.body.classList.add('modal-open');
}


// Close weekly schedule popup
function closeWeeklyModal(){

  $('weeklyModal').classList.add('hidden');

  document.body.classList.remove('modal-open');
}


async function downloadWeeklyPNG(){

  const panel=$('weeklyModal').querySelector('.weekly-modal-panel');

  if(typeof html2canvas==='undefined'){
    alert('PNG export library could not load. Please check your internet connection.');
    return;
  }

  const button=$('downloadWeeklyPng');

  try{

    button.disabled=true;
    button.textContent='Generating PNG...';

    const canvas=await html2canvas(panel,{

      scale:2,
      backgroundColor:'#ffffff',
      useCORS:true,
      logging:false,

      width:panel.scrollWidth,
      height:panel.scrollHeight,

      windowWidth:panel.scrollWidth,
      windowHeight:panel.scrollHeight,

      onclone:(clonedDoc)=>{

        const clonedPanel=clonedDoc.querySelector('.weekly-modal-panel');

        if(clonedPanel){

          clonedPanel.style.maxHeight='none';
          clonedPanel.style.height='auto';
          clonedPanel.style.overflow='visible';

        }

        const tableContainer=clonedDoc.querySelector('.weekly-table-container');

        if(tableContainer){
          tableContainer.style.overflow='visible';
        }

      }

    });

    canvas.toBlob(blob=>{

      if(!blob){
        alert('Could not generate PNG image.');
        return;
      }

      const url=URL.createObjectURL(blob);

      const link=document.createElement('a');

      const title=$('weeklyModalTitle').textContent
        .replace(/[^a-z0-9]+/gi,'-')
        .replace(/^-|-$/g,'');

      link.href=url;
      link.download=`${title}.png`;

      document.body.appendChild(link);
      link.click();
      link.remove();

      setTimeout(()=>URL.revokeObjectURL(url),1000);

    },'image/png');

  }
  catch(error){

    console.error('PNG Export Error:',error);

    alert('Unable to download PNG. Please try again.');

  }
  finally{

    button.disabled=false;
    button.textContent='⬇ Download PNG';

  }

}



function isLabCourse(r){
  const text=`${r.code||''} ${r.course||''}`.toLowerCase();
  return /\b(lab|laboratory|practical)\b/.test(text);
}

function normalizeDayTokens(v){
  const raw=String(v??'').toUpperCase().trim();
  if(!raw)return [];

  const aliases={
    SUN:'SU',SUNDAY:'SU',
    MON:'M',MONDAY:'M',
    TUE:'T',TUESDAY:'T',TUES:'T',
    WED:'W',WEDNESDAY:'W',
    THU:'R',THURSDAY:'R',THURS:'R',
    FRI:'F',FRIDAY:'F',
    SAT:'S',SATURDAY:'S'
  };

  const words=raw.replace(/[,/#+&]/g,' ').split(/\s+/).filter(Boolean);
  const out=[];

  for(const w of words){
    if(aliases[w]){
      out.push(aliases[w]);
      continue;
    }

    if(w==='SU'){
      out.push('SU');
      continue;
    }

    for(const ch of w.replace(/[^MTWRFSU]/g,'')){
      out.push(ch);
    }
  }

  return [...new Set(out)];
}

/*
  AIUB non-lab pattern:
  Sunday + Tuesday share the same repeated class-time slot.
  Monday + Wednesday share the same repeated class-time slot.
  Lab/practical classes keep their exact day identity.
*/

function logicalDays(r,meeting=r){

  const days=normalizeDayTokens(meeting.days);

  const type=String(meeting.type||'')
    .trim()
    .toLowerCase();

  // Use the actual meeting type from Excel
  const lab=type
    ? /\b(lab|laboratory|practical)\b/.test(type)
    : isLabCourse(r);

  // Lab classes retain their actual weekdays
  if(lab) return days;

  // Preserve existing AIUB theory scheduling logic
  const out=[];

  for(const d of days){

    if(d==='SU' || d==='T'){
      out.push('SUN_TUE');
    }
    else if(d==='M' || d==='W'){
      out.push('MON_WED');
    }
    else{
      out.push(d);
    }

  }

  return [...new Set(out)];
}


function meetings(r){

  // Use all meetings belonging to this section
  const entries=Array.isArray(r.meetings) && r.meetings.length
    ? r.meetings
    : [r];

  const result=[];

  for(const entry of entries){

    const range=timeRange(entry.time);
    const days=logicalDays(r,entry);

    if(!range || range[1]<=range[0] || !days.length){
      continue;
    }

    days.forEach(day=>{

      result.push({
        day,
        start:range[0],
        end:range[1]
      });

    });

  }

  return result;
}


function conflict(a,b){
  const ma=meetings(a);
  const mb=meetings(b);

  if(!ma.length || !mb.length) return true;

  return ma.some(x=>mb.some(y=>
    x.day===y.day &&
    x.start<y.end &&
    y.start<x.end
  ));
}

/*
  Total idle minutes between consecutive classes.
  Lower total gap = better schedule.
*/
function scheduleGap(schedule){
  const byDay={};

  for(const r of schedule){
    for(const m of meetings(r)){
      (byDay[m.day]??=[]).push({
        start:m.start,
        end:m.end
      });
    }
  }

  let total=0;
  let largest=0;

  for(const list of Object.values(byDay)){
    list.sort((a,b)=>a.start-b.start || a.end-b.end);

    for(let i=1;i<list.length;i++){
      const gap=Math.max(0,list[i].start-list[i-1].end);
      total+=gap;
      if(gap>largest) largest=gap;
    }
  }

  return {total,largest};
}

function sortSchedulesByGap(list){
  list.sort((a,b)=>{
    const ga=scheduleGap(a);
    const gb=scheduleGap(b);

    return ga.total-gb.total ||
           ga.largest-gb.largest ||
           signature(a).localeCompare(signature(b));
  });

  return list;
}

function signature(s){

  return s.map(x=>
    `${x.classId||''}|${x.code||x.course}|${x.section}`
  ).sort().join('||');

}
function shuffle(a){
  for(let i=a.length-1;i>0;i--){
    const j=Math.floor(Math.random()*(i+1));
    [a[i],a[j]]=[a[j],a[i]];
  }
}


function eligibleSectionsForGroup(groupKey){
  const min=Number($('minCount').value||0);
  const max=Number($('maxCount').value||Infinity);
  return (groups[groupKey]||[]).filter(r=>isOpen(r)&&r.count>=min&&r.count<=max);
}

function syncFixedSelections(){
  fixedSelections=[];
  document.querySelectorAll('.fixed-row').forEach(row=>{
    const groupKey=row.querySelector('.fixed-course').value;
    const idx=row.querySelector('.fixed-section').value;
    if(groupKey && idx!==''){
      const sec=(groups[groupKey]||[])[Number(idx)];
      if(sec)fixedSelections.push({groupKey,section:sec});
    }
  });
  if($('useFixed').checked){
    const wanted=Math.max(1,Math.min(10,Number($('fixedCount').value||1)));
    $('fixedMessage').textContent=`${fixedSelections.length}/${wanted} constant course(s) fully selected.`;
  }
}

function renderFixedRows(){
  const enabled=$('useFixed').checked;
  $('fixedSettings').classList.toggle('hidden',!enabled);
  if(!enabled){
    fixedSelections=[];
    $('fixedRows').innerHTML='';
    $('fixedMessage').textContent='';
    return;
  }

  const count=Math.max(1,Math.min(10,Number($('fixedCount').value||1)));
  const previous=fixedSelections.map(x=>({
    groupKey:x.groupKey,
    sectionSig:signature([x.section])
  }));
  const box=$('fixedRows');
  box.innerHTML='';

  const courses=Object.entries(groups).map(([key,rs])=>({
    key,label:rs[0].code||rs[0].course,title:rs[0].course||''
  })).sort((a,b)=>a.label.localeCompare(b.label));

  for(let i=0;i<count;i++){
    const row=document.createElement('div');
    row.className='fixed-row';
    row.innerHTML=`
      <label>Constant Course ${i+1}
        <select class="fixed-course">
          <option value="">Select course...</option>
          ${courses.map(c=>`<option value="${esc(c.key)}">${esc(c.label)}${c.title&&c.title!==c.label?' — '+esc(c.title):''}</option>`).join('')}
        </select>
      </label>
      <label>Fixed Section
        <select class="fixed-section" disabled>
          <option value="">Select section...</option>
        </select>
      </label>
      <div class="fixed-status"></div>`;

    const course=row.querySelector('.fixed-course');
    const section=row.querySelector('.fixed-section');
    const status=row.querySelector('.fixed-status');

    function populate(preferredSig=''){
      const groupKey=course.value;
      section.innerHTML='<option value="">Select section...</option>';
      if(!groupKey){
        section.disabled=true;
        status.textContent='';
        return;
      }

      const eligible=eligibleSectionsForGroup(groupKey);
      section.disabled=!eligible.length;
      status.textContent=eligible.length?`${eligible.length} eligible section(s)`:'No eligible Open section in count range.';

      eligible.forEach(sec=>{
        const sourceIndex=(groups[groupKey]||[]).indexOf(sec);
        const opt=document.createElement('option');
        opt.value=String(sourceIndex);
        opt.dataset.sig=signature([sec]);
        opt.textContent=`Section ${sec.section||'-'} • Count ${sec.count} • ${sec.days||'-'} • ${displayTime(sec.time)}`;
        section.appendChild(opt);
      });

      if(preferredSig){
        const match=[...section.options].find(o=>o.dataset.sig===preferredSig);
        if(match)section.value=match.value;
      }
    }

    course.addEventListener('change',()=>{populate();syncFixedSelections();});
    section.addEventListener('change',syncFixedSelections);
    box.appendChild(row);

    if(previous[i]&&groups[previous[i].groupKey]){
      course.value=previous[i].groupKey;
      populate(previous[i].sectionSig);
    }
  }
  syncFixedSelections();
}

function validateFixedSelections(){
  if(!$('useFixed').checked)return {ok:true,fixed:[]};

  syncFixedSelections();
  const wanted=Math.max(1,Math.min(10,Number($('fixedCount').value||1)));

  if(fixedSelections.length!==wanted){
    return {ok:false,message:`Please choose both course and section for all ${wanted} constant course(s).`};
  }

  const seen=new Set();
  const min=Number($('minCount').value||0);
  const max=Number($('maxCount').value||Infinity);

  for(const item of fixedSelections){
    if(seen.has(item.groupKey)){
      return {ok:false,message:'The same course cannot be selected as a constant course more than once.'};
    }
    seen.add(item.groupKey);

    const sec=item.section;
    if(!isOpen(sec)||sec.count<min||sec.count>max){
      return {ok:false,message:`Fixed section ${sec.code||sec.course} [${sec.section||'-'}] is not eligible under current Open/count filters.`};
    }
  }

  for(let i=0;i<fixedSelections.length;i++){
    for(let j=i+1;j<fixedSelections.length;j++){
      if(conflict(fixedSelections[i].section,fixedSelections[j].section)){
        const a=fixedSelections[i].section,b=fixedSelections[j].section;
        return {ok:false,message:`Fixed sections clash: ${a.code||a.course} [${a.section||'-'}] and ${b.code||b.course} [${b.section||'-'}].`};
      }
    }
  }

  return {ok:true,fixed:fixedSelections.map(x=>x.section)};
}

function generate(more=false){
  const min=Number($('minCount').value||0);
  const max=Number($('maxCount').value||Infinity);
  const target=Math.min(50,Number($('n').value||10));

  if(min>max){alert('Minimum Count cannot be greater than Maximum Count.');return;}

  const fixedCheck=validateFixedSelections();
  if(!fixedCheck.ok){alert(fixedCheck.message);return;}

  const fixedSections=fixedCheck.fixed;
  const fixedGroups=new Set(fixedSelections.map(x=>x.groupKey));
  const codes=[...selected];

  if(!codes.length&&!fixedSections.length){alert('Select at least one course.');return;}

  const autoCodes=codes.filter(c=>!fixedGroups.has(c));
  const options=autoCodes.map(c=>
    groups[c].filter(r=>
      isOpen(r)&&r.count>=min&&r.count<=max&&
      !fixedSections.some(f=>conflict(f,r))
    )
  );

  const missing=options.map((arr,i)=>arr.length?null:autoCodes[i]).filter(Boolean);
  if(missing.length){
    const labels=missing.map(k=>groups[k]?.[0]?.code||groups[k]?.[0]?.course||k);
    alert(`No eligible non-clashing section is available for: ${labels.join(', ')}.\nA constant course/section may be blocking it.`);
    return;
  }

  const pref=$('pref').value;
  for(const arr of options){
    if(pref==='low')arr.sort((x,y)=>x.count-y.count);
    else if(pref==='high')arr.sort((x,y)=>y.count-x.count);
    else shuffle(arr);
  }

  const order=[...options.keys()].sort((a,b)=>options[a].length-options[b].length);
  const seen=new Set(more?generated.map(signature):[]);
  const found=[];
  let tries=0;
  const poolLimit=Math.min(1500,Math.max(target*40,150));

  function search(i,chosen){
    if(found.length>=poolLimit||tries++>600000)return;

    if(i===order.length){
      const complete=[...fixedSections,...chosen];
      const sig=signature(complete);
      if(!seen.has(sig)){seen.add(sig);found.push(complete);}
      return;
    }

    const arr=[...options[order[i]]];
    if(pref==='random')shuffle(arr);

    for(const section of arr){
      if(!chosen.some(x=>conflict(x,section))){
        chosen.push(section);
        search(i+1,chosen);
        chosen.pop();
      }
      if(found.length>=poolLimit)return;
    }
  }

  if(!autoCodes.length){
    const sig=signature(fixedSections);
    if(!seen.has(sig))found.push([...fixedSections]);
  }else{
    search(0,[]);
  }

  sortSchedulesByGap(found);
  const additions=found.slice(0,target);

  if(!more)generated=[];
  generated.push(...additions);

  const unique=new Map();
  for(const sch of generated)unique.set(signature(sch),sch);
  generated=[...unique.values()];
  sortSchedulesByGap(generated);

  $('resultCard').classList.remove('hidden');
  const fixedText=fixedSections.length?` • ${fixedSections.length} constant course(s) locked`:'';
  $('info').textContent=`${generated.length} schedule(s) • ranked by smallest class gap • Open only • ${min} ≤ Count ≤ ${max}${fixedText}`;

  const result=$('results');
  result.innerHTML='';

  if(!generated.length){
    result.innerHTML='<div class="note">No clash-free schedule found. Try changing fixed sections, count range, or selected courses.</div>';
    return;
  }

  const fixedSigs=new Set(fixedSections.map(x=>signature([x])));


    
generated.forEach((schedule,i)=>{

  const gap=scheduleGap(schedule);

  const card=document.createElement('div');
  card.className='schedule';

  card.innerHTML=
    `<div class="sh">

       <div>
         Schedule ${i+1} — Total Gap: ${gap.total} min
         ${gap.largest?` • Largest Single Gap: ${gap.largest} min`:''}
       </div>

       <button class="weekly-view-btn" type="button">
         View Weekly Table
       </button>

     </div>

     <table class="routine">

       <thead>
         <tr>
           <th>Course</th>
           <th>Section</th>
           <th>Type</th>
           <th>Day / Time</th>
           <th>Faculty</th>
           <th>Room</th>
           <th>Status</th>
           <th>Count</th>
         </tr>
       </thead>

       <tbody>
         ${schedule.map(x=>{

           const fixed=fixedSigs.has(signature([x]));

           return `<tr>

             <td>
               <b>${esc(x.code)}</b><br>
               ${esc(x.course)}
             </td>

             <td>
               <span class="pill">${esc(x.section||'-')}</span>
             </td>

             <td>
               ${fixed
                 ? '<span class="pill">Constant</span>'
                 : 'Auto'}
             </td>

            <td>
              ${displayScheduleMeetings(x)}
            </td>

             <td>${esc(x.faculty||'-')}</td>
             <td>${esc(x.room||'-')}</td>

             <td>
               <span class="pill">${esc(x.status||'Open')}</span>
             </td>

             <td>${x.count}</td>

           </tr>`;

         }).join('')}

       </tbody>
     </table>`;

  // Weekly Table Button
  card.querySelector('.weekly-view-btn')
    .addEventListener('click',()=>{

      openWeeklyModal(schedule,i+1);

    });

  result.appendChild(card);

});

  
}

$('browse').addEventListener('click',()=>{
  $('file').click();
});
$('file').addEventListener('change',e=>{
  const file=e.target.files&&e.target.files[0];
  if(file)loadFile(file);
});

const drop=$('drop');
drop.addEventListener('dragover',e=>{
  e.preventDefault();
  drop.classList.add('drag');
});
drop.addEventListener('dragleave',()=>{
  drop.classList.remove('drag');
});
drop.addEventListener('drop',e=>{
  e.preventDefault();
  drop.classList.remove('drag');
  const file=e.dataTransfer.files&&e.dataTransfer.files[0];
  if(file)loadFile(file);
});

$('useFixed').addEventListener('change',renderFixedRows);
$('fixedCount').addEventListener('input',renderFixedRows);
$('search').addEventListener('input',renderCourses);
$('minCount').addEventListener('input',()=>{renderCourses();renderFixedRows();});
$('maxCount').addEventListener('input',()=>{renderCourses();renderFixedRows();});

$('visible').addEventListener('click',()=>{
  document.querySelectorAll('#courses .course').forEach(el=>{
    el.querySelector('input').checked=true;
    selected.add(el.dataset.group);
  });
  updateSelected();
});
$('clear').addEventListener('click',()=>{
  selected.clear();
  renderCourses();
});
$('generate').addEventListener('click',()=>generate(false));
$('more').addEventListener('click',()=>generate(true));


// Close popup using close button
$('closeWeeklyModal').addEventListener('click',closeWeeklyModal);


// Close popup by clicking outside
$('weeklyModal').addEventListener('click',e=>{

  if(e.target.hasAttribute('data-close-weekly')){
    closeWeeklyModal();
  }

});


// Close popup using Escape key
document.addEventListener('keydown',e=>{

  if(
    e.key==='Escape' &&
    !$('weeklyModal').classList.contains('hidden')
  ){

    closeWeeklyModal();

  }

});

// Download Weekly Table as PNG
$('downloadWeeklyPng').addEventListener(
  'click',
  downloadWeeklyPNG
);
