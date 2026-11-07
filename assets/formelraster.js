"use strict";
const tasks=challenges;
const $=id=>document.getElementById(id);
const bestScoreKey="formelraster-best-v1";
const width=8,height=10;
let mode="ranked",task=tasks[0],selected=[0,0],cells=[],round=[],index=0,score=0,phase="ready",deadline=0,timerId=null,errors=0,usedHint=false,practiceIndex=0,creativeIndex=-1,lastResult="";
const endlessProgressKey="formelraster-endless-v1";
let endlessSolved=0,endlessTask=null,endlessNext=true;
try{const saved=Number(localStorage.getItem(endlessProgressKey));if(Number.isSafeInteger(saved)&&saved>=0)endlessSolved=saved}catch{}
let record=0;
try{record=Math.max(Number(localStorage.getItem(bestScoreKey))||0,Number(localStorage.getItem("excel-ranked-best-v1"))||0)}catch{}
const nameOf=(r,c)=>String.fromCharCode(65+c)+(r+1);
const asNumber=value=>typeof value==="boolean"?Number(value):value===""?0:Number(String(value).trim().replace(",","."));
const normalized=value=>String(value).trim().toLocaleLowerCase("de-DE");
const equal=(a,b)=>{const left=asNumber(a),right=asNumber(b);return String(a).trim()!==""&&String(b).trim()!==""&&Number.isFinite(left)&&Number.isFinite(right)?Math.abs(left-right)<0.000001:normalized(a)===normalized(b)};
const functionNames={SUMME:"SUM",MITTELWERT:"AVERAGE",ANZAHL:"COUNT",WENN:"IF",RUNDEN:"ROUND","ZÄHLENWENN":"COUNTIF",SUMMEWENN:"SUMIF",XVERWEIS:"XLOOKUP"};
function callFunction(name,args){
 const values=args.flat(Infinity), numbers=values.filter(v=>v!==""&&Number.isFinite(asNumber(v))).map(asNumber);
 const [first,second,third]=args, list=v=>Array.isArray(v)?v:[v];
 switch(functionNames[name]||name){
 case "SUM":return numbers.reduce((a,b)=>a+b,0);
 case "AVERAGE":return numbers.reduce((a,b)=>a+b,0)/numbers.length;
 case "MIN":return numbers.length?Math.min(...numbers):0;
 case "MAX":return numbers.length?Math.max(...numbers):0;
 case "COUNT":return numbers.length;
 case "ROUND":{const factor=10**asNumber(second);return Math.sign(asNumber(first))*Math.round(Math.abs(asNumber(first))*factor)/factor}
 case "IF":return first?second:third;
 case "COUNTIF":return list(first).filter(v=>equal(v,second)).length;
 case "SUMIF":{const totals=list(third??first);return list(first).reduce((sum,v,i)=>sum+(equal(v,second)?asNumber(totals[i]):0),0)}
 case "XLOOKUP":{const i=list(second).findIndex(v=>equal(v,first));if(i<0)throw Error("#NV");return list(third)[i]}
 default:throw Error("#NAME?");
 }
}
function formulaValue(text,stack=new Set()){
 const tokens=[];let rest=text.slice(1);
 while(rest.trim()){
  rest=rest.trimStart();
  const match=rest.match(/^(?:"(?:[^"]|"")*"|\d+(?:[.,]\d+)?|#BEZUG!|\$?[A-H]\$?(?:10|[1-9])(?![\p{L}\d_])|[\p{L}_]+\d*|>=|<=|<>|[+\-*/^%():;=><,])/u);
  if(!match)throw Error("#FORMEL!");tokens.push(match[0]);rest=rest.slice(match[0].length);
 }
 let i=0;
 const take=value=>tokens[i]===value?(i++,true):false;
 const requireToken=value=>{if(!take(value))throw Error("#FORMEL!")};
 function ref(token){const m=/^\$?([A-H])\$?(10|[1-9])$/i.exec(token??"");if(!m)throw Error("#BEZUG!");return [+m[2]-1,m[1].toUpperCase().charCodeAt(0)-65]}
 function atom(){
  if(take("(")){const value=compare();requireToken(")");return value}
  const token=tokens[i++];if(!token)throw Error("#FORMEL!");
  if(token.startsWith('"'))return token.slice(1,-1).replaceAll('""','"');
  if(/^\d/.test(token))return asNumber(token);
  const upper=token.toUpperCase();
  if(take("(")){
   const args=[];
   if(!take(")")){do{args.push(compare())}while(take(";")||take(","));requireToken(")")}
   return callFunction(upper,args);
  }
  if(["TRUE","WAHR"].includes(upper))return true;
  if(["FALSE","FALSCH"].includes(upper))return false;
  const [r,c]=ref(upper);
  if(!take(":"))return cellValue(r,c,stack);
  const [endR,endC]=ref(tokens[i++]), values=[];
  for(let rr=Math.min(r,endR);rr<=Math.max(r,endR);rr++)for(let cc=Math.min(c,endC);cc<=Math.max(c,endC);cc++)values.push(cellValue(rr,cc,stack));
  return values;
 }
 function power(){let value=atom();while(take("%"))value=asNumber(value)/100;return take("^")?asNumber(value)**asNumber(unary()):value}
 function unary(){if(take("-"))return -asNumber(unary());if(take("+"))return asNumber(unary());return power()}
 function product(){let value=unary();while(["*","/"].includes(tokens[i])){const op=tokens[i++],rhs=asNumber(unary());value=op==="*"?asNumber(value)*rhs:asNumber(value)/rhs}return value}
 function sum(){let value=product();while(["+","-"].includes(tokens[i])){const op=tokens[i++],rhs=asNumber(product());value=op==="+"?asNumber(value)+rhs:asNumber(value)-rhs}return value}
 function compare(){const a=sum(),op=tokens[i];if(!["=","<>",">","<",">=","<="].includes(op))return a;i++;const b=sum();if(op==="=")return equal(a,b);if(op==="<>")return !equal(a,b);const x=asNumber(a),y=asNumber(b);return op===">"?x>y:op==="<"?x<y:op===">="?x>=y:x<=y}
 const value=compare();
 if(i!==tokens.length||Array.isArray(value))throw Error("#FORMEL!");
 if(typeof value==="number"&&!Number.isFinite(value))throw Error("#ZAHL!");
 return value;
}
function cellValue(r,c,stack=new Set()){
 const value=cells[r]?.[c]??"", key=nameOf(r,c);
 if(!value.startsWith("="))return value;
 if(stack.has(key))throw Error("#KREIS!");
 stack.add(key);try{return formulaValue(value,stack)}finally{stack.delete(key)}
}
const inputs=[];
let selectionEnd=[0,0], gesture=null, referenceArea=null, formulaEditor=null, touchRange=false;
let lastCopy=null;
const sameCell=(a,b)=>a[0]===b[0]&&a[1]===b[1];
function rect(a,b=a){return [Math.min(a[0],b[0]),Math.min(a[1],b[1]),Math.max(a[0],b[0]),Math.max(a[1],b[1])]}
function contains(area,r,c){return !!area&&r>=area[0]&&r<=area[2]&&c>=area[1]&&c<=area[3]}
function rangeName(area){const a=nameOf(area[0],area[1]),b=nameOf(area[2],area[3]);return a===b?a:a+":"+b}
function changeReferences(value,update){
 if(!value.startsWith("="))return value;
 return value.replace(/"(?:[^"]|"")*"|(?<![\p{L}\d_])(\$?)([A-H])(\$?)(10|[1-9])(?![\p{L}\d_])/giu,(match,colLock,col,rowLock,row)=>{
  if(!col)return match;
  const r=Number(row)-1,c=col.toUpperCase().charCodeAt(0)-65;
  const [newR,newC]=update(r,c,!!rowLock,!!colLock);
  if(newR===r&&newC===c)return match;
  if(!contains([0,0,height-1,width-1],newR,newC))return "#BEZUG!";
  return colLock+String.fromCharCode(65+newC)+rowLock+(newR+1);
 });
}
function shiftedFormula(value,dr,dc){return changeReferences(value,(r,c,fixedRow,fixedCol)=>[r+(fixedRow?0:dr),c+(fixedCol?0:dc)])}
function fillArea(source,target,vertical){
 const original=cells.map(row=>row.slice());
 const rowCount=source[2]-source[0]+1,colCount=source[3]-source[1]+1;
 const period=vertical?rowCount:colCount;
 for(let r=target[0];r<=target[2];r++)for(let c=target[1];c<=target[3];c++){
  if(contains(source,r,c))continue;
  const offset=vertical?r-source[0]:c-source[1];
  const wrapped=((offset%period)+period)%period;
  const rr=vertical?source[0]+wrapped:r,cc=vertical?c:source[1]+wrapped;
  const series=Array.from({length:period},(_,i)=>vertical?original[source[0]+i][c]:original[r][source[1]+i]);
  const numbers=series.map(asNumber),step=numbers[1]-numbers[0];
  const isSeries=period>1&&series.every((v,i)=>v.trim()!==""&&Number.isFinite(numbers[i])&&Math.abs(numbers[i]-(numbers[0]+i*step))<1e-8);
  cells[r][c]=isSeries?String(Number((numbers[0]+offset*step).toPrecision(12))):shiftedFormula(original[rr][cc],r-rr,c-cc);
 }
}
function moveArea(source,target){
 const dr=target[0]-source[0],dc=target[1]-source[1];
 if(!dr&&!dc)return;
 const original=cells.map(row=>row.slice());
 for(let r=0;r<height;r++)for(let c=0;c<width;c++){
  const value=contains(target,r,c)?original[r-dr][c-dc]:contains(source,r,c)?"":original[r][c];
  cells[r][c]=changeReferences(value,(rr,cc)=>contains(source,rr,cc)?[rr+dr,cc+dc]:[rr,cc]);
 }
}
function paintSelection(){
 const area=rect(selected,selectionEnd), target=gesture?.target;
 $("address").textContent=rangeName(area);
 for(let r=0;r<height;r++)for(let c=0;c<width;c++){
  const td=inputs[r][c].parentElement,inside=contains(area,r,c);
  td.classList.toggle("selected",sameCell(selected,[r,c]));
  td.classList.toggle("in-range",inside);
  td.classList.toggle("range-top",inside&&r===area[0]);
  td.classList.toggle("range-bottom",inside&&r===area[2]);
  td.classList.toggle("range-left",inside&&c===area[1]);
  td.classList.toggle("range-right",inside&&c===area[3]);
  td.classList.toggle("drag-target",contains(target,r,c)&&!contains(gesture?.source,r,c));
  td.classList.toggle("ref-target",contains(referenceArea,r,c));
 }
 for(const kind of ["move","fill"]){
  const handle=$(kind+"Handle"),r=kind==="move"?area[0]:area[2],c=kind==="move"?area[1]:area[3];
  const corner=inputs[r][c].parentElement;
  if(handle.parentElement!==corner)corner.append(handle);
  handle.hidden=inputs[r][c].disabled||!!gesture||!!referenceArea;
 }
}
function startHandle(event,kind){
 if(event.button!==0||$("formula").disabled)return;
 event.preventDefault();event.stopPropagation();
 document.activeElement.blur();
 const source=rect(selected,selectionEnd);
 gesture={kind,source,target:source,pointer:event.pointerId};
 $("grid").classList.add("is-dragging");paintSelection();
}
function startSelection(event,r,c){
 if(event.button!==0||$("formula").disabled)return;
 const editor=formulaEditor;
 if(editor&&document.activeElement===editor&&editor.value.startsWith("=")){
  const start=editor.selectionStart,end=editor.selectionEnd;
  if(start!==end||/[=(;,+:*/^<>-]\s*$/.test(editor.value.slice(0,start))){
   event.preventDefault();
   gesture={kind:"reference",pointer:event.pointerId,anchor:[r,c],editor,start,end};
   referenceArea=rect([r,c]);paintSelection();return;
  }
 }
 if(touchRange){event.preventDefault();selectionEnd=[r,c];setTouchRange(false);paintSelection();message("Bereich "+rangeName(rect(selected,selectionEnd))+" markiert.");return}
 if(event.shiftKey){event.preventDefault();selectionEnd=[r,c];paintSelection()}
 else selectCell(r,c);
 gesture={kind:"select",pointer:event.pointerId,anchor:selected,x:event.clientX,y:event.clientY};
}
function movePointer(event){
 if(!gesture||gesture.pointer!==event.pointerId)return;
 const td=document.elementFromPoint(event.clientX,event.clientY)?.closest("td[data-cell]");
 if(!td)return;
 const [r,c]=td.dataset.cell.split(",").map(Number);
 const {kind,source,anchor}=gesture;
 if(kind==="select"){
  if(Math.hypot(event.clientX-gesture.x,event.clientY-gesture.y)<4)return;
  $("grid").classList.add("is-dragging");selectionEnd=[r,c];
  const active=document.activeElement;
  if(active instanceof HTMLInputElement&&$("grid").contains(active))active.setSelectionRange(0,0);
 }else if(kind==="reference"){
  referenceArea=rect(anchor,[r,c]);
 }else if(kind==="move"){
  const rows=source[2]-source[0],cols=source[3]-source[1];
  const top=Math.min(r,height-rows-1),left=Math.min(c,width-cols-1);
  gesture.target=[top,left,top+rows,left+cols];
 }else{
  const dr=r<source[0]?r-source[0]:Math.max(0,r-source[2]);
  const dc=c<source[1]?c-source[1]:Math.max(0,c-source[3]);
  gesture.vertical=Math.abs(dr)>=Math.abs(dc);
  gesture.target=gesture.vertical?[Math.min(r,source[0]),source[1],Math.max(r,source[2]),source[3]]:[source[0],Math.min(c,source[1]),source[2],Math.max(c,source[3])];
 }
 paintSelection();
}
function finishPointer(event,cancel=false){
 if(!gesture||gesture.pointer!==event.pointerId)return;
 const {kind,source,target,vertical,editor,start,end}=gesture;
 const area=referenceArea;
 gesture=null;referenceArea=null;$("grid").classList.remove("is-dragging");
 if(!cancel&&kind==="reference"&&area){
  editor.setRangeText(rangeName(area),start,end,"end");
 }else if(!cancel&&(kind==="move"||kind==="fill")){
  if(kind==="move")moveArea(source,target);else fillArea(source,target,vertical);
  selected=[target[0],target[1]];selectionEnd=[target[2],target[3]];
  $("formula").value=cells[selected[0]][selected[1]];renderGrid();
 }
 paintSelection();
}
function copySelection(cut=false){
 const area=rect(selected,selectionEnd);
 const data=cells.slice(area[0],area[2]+1).map(row=>row.slice(area[1],area[3]+1));
 const plain=data.map(row=>row.join("\t")).join("\n");
 lastCopy={plain,area,data,cut,grid:cells};return plain;
}
function clearSelection(){
 const [top,left,bottom,right]=rect(selected,selectionEnd);
 for(let r=top;r<=bottom;r++)for(let c=left;c<=right;c++)cells[r][c]="";
 const active=document.activeElement;
 if($("grid").contains(active)&&active.matches("input"))active.value="";
 $("formula").value="";renderGrid();
}
function clipboardCopy(event,cut=false){
 const active=document.activeElement;
 if(sameCell(selected,selectionEnd)&&active.selectionStart!==active.selectionEnd)return;
 if(!event.clipboardData)return;
 event.clipboardData.setData("text/plain",copySelection(cut));event.preventDefault();
}
function pasteSelection(plain){
 const matching=lastCopy?.plain===plain?lastCopy:null;
 const [top,left]=selected;
 if(matching?.cut&&matching.grid===cells){
  const rows=matching.data.length,cols=matching.data[0].length;
  if(top+rows>height||left+cols>width){message("Der ausgeschnittene Bereich passt hier nicht hinein.","bad");return}
  moveArea(matching.area,[top,left,top+rows-1,left+cols-1]);lastCopy=null;
 }else{
  const data=matching?.data??plain.replace(/\r\n?/g,"\n").replace(/\n$/,"").split("\n").map(row=>row.split("\t"));
  for(let r=0;r<data.length&&top+r<height;r++)for(let c=0;c<data[r].length&&left+c<width;c++){
   cells[top+r][left+c]=matching?shiftedFormula(data[r][c],top-matching.area[0],left-matching.area[1]):data[r][c];
  }
 }
 const active=document.activeElement;
 if($("grid").contains(active)&&active.matches("input"))active.value=cells[top][left];
 renderGrid();$("formula").value=cells[top][left];
}
function clipboardPaste(event){
 const plain=event.clipboardData?.getData("text/plain");
 if(plain===undefined||(!plain&&lastCopy?.plain!==""))return;
 const active=document.activeElement;
 if(!/[\t\r\n]/.test(plain)&&lastCopy?.plain!==plain&&active.selectionStart!==active.selectionEnd)return;
 event.preventDefault();pasteSelection(plain);
}
let menuReturn=null;
function openCellMenu(event,r,c){
 event.preventDefault();if($("formula").disabled)return;
 document.activeElement.blur();
 if(!contains(rect(selected,selectionEnd),r,c))selectCell(r,c);
 menuReturn=inputs[selected[0]][selected[1]];
 const menu=$("cellMenu");$("menuAddress").textContent=rangeName(rect(selected,selectionEnd));
 menu.hidden=false;
 const target=event.currentTarget.getBoundingClientRect();
 const x=event.clientX||target.left,y=event.clientY||target.bottom;
 menu.style.left=Math.max(6,Math.min(x,innerWidth-menu.offsetWidth-6))+"px";
 menu.style.top=Math.max(6,Math.min(y,innerHeight-menu.offsetHeight-6))+"px";
 menu.querySelector("button").focus();
}
function closeCellMenu(restore=false){
 $("cellMenu").hidden=true;
 if(restore&&menuReturn){const [r,c]=selectionEnd;menuReturn.focus();selectionEnd=[r,c];paintSelection()}
}
function restructure(axis,insert){
 const limit=axis===0?height:width,index=selected[axis],old=cells.map(row=>row.slice());
 if(insert&&old.some((row,r)=>row.some((value,c)=>value!==""&&[r,c][axis]===limit-1))){
  message("Am Tabellenrand ist kein Platz. Leere zuerst die letzte Zeile oder Spalte.","bad");return;
 }
 for(let r=0;r<height;r++)for(let c=0;c<width;c++){
  const previous=[r,c];
  if(insert&&previous[axis]===index){cells[r][c]="";continue}
  if(previous[axis]>=index)previous[axis]+=insert?-1:1;
  const value=old[previous[0]]?.[previous[1]]??"";
  cells[r][c]=changeReferences(value,(rr,cc)=>{
   const ref=[rr,cc];
   if(!insert&&ref[axis]===index)return [-1,-1];
   if(ref[axis]>=index)ref[axis]+=insert?1:-1;
   return ref;
  });
 }
 lastCopy=null;renderGrid();selectCell(...selected);
}
async function menuAction(action){
 closeCellMenu();
 if(action==="copy"||action==="cut"){
  const plain=copySelection(action==="cut");
  try{await navigator.clipboard.writeText(plain)}catch{}
  message(action==="cut"?"Bereich zum Ausschneiden vorgemerkt. Wähle das Ziel und füge ihn ein.":"Bereich kopiert.","good");
 }else if(action==="paste"){
  if(lastCopy){pasteSelection(lastCopy.plain);return}
  try{
   const plain=await navigator.clipboard.readText();
   if(plain)pasteSelection(plain);else message("Die Zwischenablage ist leer.");
  }catch{menuReturn?.focus();message("Drücke Strg+V, um aus der Zwischenablage einzufügen.")}
 }else if(action==="clear")clearSelection();
 else if(action==="sum")autoSum();
 else restructure(action.endsWith("row")?0:1,action.startsWith("insert"));
}
function autoSum(){
 const area=rect(selected,selectionEnd),[r,c]=selected;
 let destination,source;
 if(!sameCell(selected,selectionEnd)){
  if(area[2]+1>=height){message("Unter der Auswahl ist keine freie Zeile für AutoSumme.","bad");return}
  destination=[area[2]+1,area[1]];source=area;
 }else{
  if(!r){message("Markiere einen Bereich oder wähle eine Zelle unter Zahlen.","bad");return}
  let first=r-1;
  while(first>0){try{if(cellValue(first-1,c)===""||!Number.isFinite(asNumber(cellValue(first-1,c))))break}catch{break}first--}
  destination=[r,c];source=[first,c,r-1,c];
 }
 const [rr,cc]=destination;
 if(cells[rr][cc]!==""){message("Wähle eine leere Ergebniszelle für AutoSumme.","bad");return}
 cells[rr][cc]="=SUMME("+rangeName(source)+")";renderGrid();selectCell(rr,cc);signal("cell",inputs[rr][cc].parentElement);
}

let cellEdit=null,cellBeforeEdit="";
function beginCellEdit(input,r,c){
 if(cellEdit!==input)cellBeforeEdit=cells[r][c];
 cellEdit=input;input.setSelectionRange(input.value.length,input.value.length);
}
function focusGridCell(r,c,extend=false){
 const anchor=selected.slice();
 const row=Math.max(0,Math.min(height-1,r)),column=Math.max(0,Math.min(width-1,c));
 inputs[row][column].focus({preventScroll:true});
 inputs[row][column].scrollIntoView({block:"nearest",inline:"nearest"});
 if(extend){
  selected=anchor;selectionEnd=[row,column];
  $("formula").value=cells[anchor[0]][anchor[1]];paintSelection();
 }else selectCell(row,column);
}
function gridKey(event,r,c){
 const input=event.currentTarget, key=event.key;
 if(key==="ContextMenu"||(event.shiftKey&&key==="F10")){openCellMenu(event,r,c);return}
 if(key==="F2"){event.preventDefault();beginCellEdit(input,r,c);return}
 if(key==="Escape"){
  event.preventDefault();
  if(cellEdit===input){
   cells[r][c]=cellBeforeEdit;input.value=cellBeforeEdit;cellEdit=null;renderGrid();
  }
  selectCell(r,c);return;
 }
 if(key==="Enter"||key==="Tab"){
  event.preventDefault();input.blur();cellEdit=null;
  const step=event.shiftKey?-1:1;
  if(key==="Enter")focusGridCell(r+step,c);
  else{
   const flat=r*width+c+step;
   if(flat<0||flat>=height*width){$(step<0?"formula":$("answerForm").hidden?"export":"answer").focus();return}
   focusGridCell(Math.floor(flat/width),flat%width);
  }
  return;
 }
 const arrows={ArrowUp:[-1,0],ArrowDown:[1,0],ArrowLeft:[0,-1],ArrowRight:[0,1]};
 if(Object.hasOwn(arrows,key)&&cellEdit!==input){
  event.preventDefault();
  const [dr,dc]=arrows[key],limit=event.ctrlKey||event.metaKey;
  focusGridCell(r+dr*(limit?height:1),c+dc*(limit?width:1),event.shiftKey);return;
 }
 if(key==="Delete"&&(!sameCell(selected,selectionEnd)||cellEdit!==input)){
  event.preventDefault();clearSelection();cellBeforeEdit=cells[r][c];return;
 }
 if(!event.ctrlKey&&!event.metaKey&&!event.altKey&&key.length===1&&cellEdit!==input){
  beginCellEdit(input,r,c);input.select();
 }
}
function makeGrid(){
 const table=document.createElement("table"),top=table.createTHead().insertRow();
 for(let c=-1;c<width;c++){const th=document.createElement("th");th.scope="col";th.textContent=c<0?"#":String.fromCharCode(65+c);top.append(th)}
 const body=table.createTBody();
 for(let r=0;r<height;r++){const row=body.insertRow(),th=document.createElement("th");th.scope="row";th.textContent=r+1;row.append(th);inputs[r]=[];
  for(let c=0;c<width;c++){const td=row.insertCell(),input=document.createElement("input");td.dataset.cell=r+","+c;td.addEventListener("pointerdown",event=>startSelection(event,r,c));td.addEventListener("contextmenu",event=>openCellMenu(event,r,c));input.setAttribute("aria-label",nameOf(r,c));input.spellcheck=false;input.autocomplete="off";td.append(input);inputs[r][c]=input;
   input.addEventListener("focus",()=>{if(gesture?.kind!=="select")selectCell(r,c);input.value=cells[r][c];cellEdit=null;cellBeforeEdit=cells[r][c]});
   input.addEventListener("input",()=>{cellEdit=input;cells[r][c]=input.value;$("formula").value=input.value;renderGrid()});
   input.addEventListener("blur",()=>{if(!input.disabled){cells[r][c]=input.value;renderGrid()}});
   input.addEventListener("keydown",event=>gridKey(event,r,c));
   input.addEventListener("dblclick",()=>beginCellEdit(input,r,c));

  }
 }
 $("grid").append(table);
 for(const kind of ["move","fill"]){
  const handle=document.createElement("span");handle.id=kind+"Handle";handle.className="cell-handle "+kind;
  handle.title=kind==="move"?"Bereich verschieben":"Ausfüllen";
  handle.addEventListener("pointerdown",event=>startHandle(event,kind));$("grid").append(handle);
 }
 $("grid").addEventListener("copy",clipboardCopy);$("grid").addEventListener("cut",event=>clipboardCopy(event,true));$("grid").addEventListener("paste",clipboardPaste);
 document.addEventListener("pointermove",movePointer);
 document.addEventListener("pointerup",event=>finishPointer(event));
 document.addEventListener("pointercancel",event=>finishPointer(event,true));
}
function selectCell(r,c){selected=[r,c];selectionEnd=[r,c];$("formula").value=cells[r][c];paintSelection()}
function renderGrid(){
 for(let r=0;r<height;r++)for(let c=0;c<width;c++){
  const input=inputs[r][c],td=input.parentElement;td.classList.toggle("header-cell",!!task&&r===0&&c<task.rows[0].length);
  if(input===document.activeElement)continue;
  try{input.value=cellValue(r,c);td.classList.remove("error")}catch(error){input.value=error.message;td.classList.add("error")}
 }
}
function loadGrid(){
 closeCellMenu();setTouchRange(false);
 cells=Array.from({length:height},(_,r)=>Array.from({length:width},(_,c)=>String(task?.rows[r]?.[c]??"")));
 renderGrid();selectCell(0,0);$("answer").value="";
}
function message(text,tone=""){$("status").textContent=text;$("status").dataset.tone=tone}
function stopClock(){clearInterval(timerId);timerId=null}
function secondsLeft(){return Math.max(0,Math.ceil((deadline-Date.now())/1000))}
function tick(){
 const left=secondsLeft();$("timer").textContent=left+" s";$("timer").classList.toggle("low",left<=10);
 if(!left&&phase==="playing"&&mode==="ranked")finishTask(false);
}
function sync(){
 const ranked=mode==="ranked",endless=mode==="endless",creative=mode==="creative",playing=phase==="playing";
 $("record").textContent=record+" Pkt.";$("score").textContent=endless?endlessSolved:ranked?score:"—";
 $("metricLabel").textContent=endless?"Stufe":"Zeit";$("scoreLabel").textContent=endless?"Gelöst":"Punkte";
 if(endless){$("timer").textContent=task?.level??endlessSolved+1;$("timer").classList.remove("low")}
 $("challenge").hidden=mode!=="practice";$("creativeProject").hidden=!creative;$("export").hidden=!creative;$("stats").hidden=!ranked&&!endless;
 $("answerForm").hidden=creative&&!task;$("answer").disabled=!playing;$("submit").disabled=!playing;
 $("hint").hidden=creative&&!task;$("hint").disabled=!playing||usedHint;$("reset").disabled=ranked&&!playing;
 $("formula").disabled=ranked&&!playing;for(const row of inputs)for(const input of row)input.disabled=ranked&&!playing;
 for(const id of ["autoSum","touchRange","touchMenu"])$(id).disabled=$("formula").disabled;paintSelection();
 $("start").hidden=creative||playing;
 $("start").textContent=phase==="result"?(ranked&&index===round.length-1?"Ergebnis ansehen":"Nächste Aufgabe"):phase==="done"?"Neue Runde":phase==="solved"?"Nächste Aufgabe":"Runde starten";
 $("sheetName").textContent=creative?"Freies Arbeitsblatt":"Arbeitsblatt · "+(task?.tier??"Ranked");
 if(creative&&task){$("stage").textContent="Kreativ · Projekt";$("taskTitle").textContent=task.title;$("brief").textContent=task.brief;return}
 if(creative){$("stage").textContent="Kreativ · ohne Uhr";$("taskTitle").textContent="Dein Tabellenlabor";$("brief").textContent="Baue deine eigene Tabelle, experimentiere mit Formeln und exportiere die Ergebnisse als CSV.";return}
 if(ranked&&(phase==="ready"||phase==="done")){
  $("stage").textContent="Ranked · Bronze bis Gold";$("taskTitle").textContent=phase==="done"?"Runde beendet":"Bereit für Ranked?";
  $("brief").textContent=phase==="done"?lastResult:"Drei Aufgaben, ansteigende Schwierigkeit. Für richtige Antworten und Restzeit gibt es Punkte.";
  $("timer").textContent="—";$("timer").classList.remove("low");return;
 }
 if(endless){$("stage").textContent="Unendlich · "+task.tier;$("taskTitle").textContent=task.title;$("brief").textContent=task.brief;return}
 $("stage").textContent=ranked?task.tier+" · Aufgabe "+(index+1)+" / "+round.length:task.tier+" · Training";
 $("taskTitle").textContent=task.title;$("brief").textContent=task.brief;
}
function setupTask(next){
 task=next;phase="playing";errors=0;usedHint=false;$("hintText").textContent="";message("");loadGrid();
 if(mode==="ranked"){deadline=Date.now()+task.time*1000;stopClock();timerId=setInterval(tick,150);tick()}
 sync();signal("task",$("taskTitle"));
}
function nextEndlessTask(){
 if(endlessNext){const level=endlessSolved+1;endlessTask={...generateEndlessTask(level),level};endlessNext=false}
 setupTask(endlessTask);
}
function switchMode(next){
 stopClock();mode=next;phase="ready";task=null;score=0;usedHint=false;$("hintText").textContent="";message("");
 document.querySelectorAll(".mode").forEach(btn=>btn.setAttribute("aria-pressed",String(btn.dataset.mode===mode)));
 if(mode==="endless"){nextEndlessTask();return}
 if(mode==="creative"&&creativeIndex>=0){setupTask(creativeProjects[creativeIndex]);return}
 if(mode==="practice"){setupTask(tasks[practiceIndex]);return}
 loadGrid();sync();
}
function finishTask(correct){
 const time=secondsLeft();stopClock();
 if(mode==="endless"){
  endlessSolved=task.level;endlessNext=true;phase="solved";
  try{localStorage.setItem(endlessProgressKey,String(endlessSolved))}catch{}
  message("Richtig! Stufe "+endlessSolved+" geschafft.","good");sync();return;
 }
 if(mode==="creative"){phase="solved";message("Projektziel erreicht. Du kannst die Tabelle weiter bearbeiten und als CSV speichern.","good");sync();return}
 if(mode==="practice"){phase="solved";message("Richtig! Weiter zur nächsten Aufgabe.","good");sync();return}
 phase="result";
 if(correct){const earned=Math.max(50,300+time*5-errors*35-(usedHint?70:0));score+=earned;message("Richtig! +"+earned+" Punkte.","good")}
 else message("Zeit abgelaufen. Lösung: "+task.answer+".","bad");
 sync();
}
$("start").onclick=()=>{
 if(mode==="endless"){nextEndlessTask();return}
 if(mode==="practice"){practiceIndex=(practiceIndex+1)%tasks.length;$("challenge").value=practiceIndex;setupTask(tasks[practiceIndex]);return}
 if(phase==="result"){
  if(index<round.length-1){setupTask(round[++index]);return}
  const best=score>record;record=Math.max(record,score);try{localStorage.setItem(bestScoreKey,String(record))}catch{}
  const rank=score>=1350?"Diamant":score>=1000?"Gold":score>=600?"Silber":"Bronze";
  lastResult=rank+" mit "+score+" Punkten"+(best?" — neuer Rekord!":".");phase="done";message("");$("hintText").textContent="";sync();return;
 }
 round=["Bronze","Silber","Gold"].map(tier=>{const pool=tasks.filter(t=>t.tier===tier);return pool[Math.floor(Math.random()*pool.length)]});
 index=0;score=0;setupTask(round[0]);
};
$("answerForm").onsubmit=e=>{
 e.preventDefault();if(phase!=="playing")return;if(mode==="ranked"&&!secondsLeft()){finishTask(false);return}
 const raw=$("answer").value.trim();if(!raw)return;
 if(!raw.startsWith("=")){message("Gib deine Antwort als Formel ein. Sie muss mit = beginnen.","bad");return}
 let value;try{value=formulaValue(raw)}catch(error){message("Formel prüfen: "+error.message,"bad");return}
 if(equal(value,task.answer)){finishTask(true);signal("success",$("status"));return}
 errors++;message("Noch nicht richtig. Du kannst es erneut versuchen.","bad");signal("retry",$("answer"));
};
$("hint").onclick=()=>{usedHint=true;$("hintText").textContent=task.hint;sync()};
$("reset").onclick=()=>{loadGrid();message("Tabelle zurückgesetzt.")}
for(const editor of [$("formula"),$("answer")])editor.addEventListener("focus",()=>{formulaEditor=editor});
$("autoSum").onclick=autoSum;
function setTouchRange(active){touchRange=active;$("touchRange").setAttribute("aria-pressed",String(active))}
$("touchRange").onclick=()=>{
 setTouchRange(!touchRange);
 message(touchRange?"Tippe auf die letzte Zelle des Bereichs.":"Bereichsauswahl beendet.");
};
$("touchMenu").onclick=event=>openCellMenu(event,...selected);
$("cellMenu").addEventListener("click",event=>{const item=event.target.closest("[data-action]");if(item)menuAction(item.dataset.action)});
$("cellMenu").addEventListener("keydown",event=>{
 const items=[...$("cellMenu").querySelectorAll("button")],index=items.indexOf(document.activeElement);
 if(event.key==="Escape"){event.preventDefault();closeCellMenu(true)}
 if(event.key==="ArrowDown"||event.key==="ArrowUp"){event.preventDefault();items[(index+(event.key==="ArrowDown"?1:-1)+items.length)%items.length].focus()}
 if(event.key==="Tab")closeCellMenu();
});
document.addEventListener("pointerdown",event=>{if(!$("cellMenu").contains(event.target))closeCellMenu()});
document.addEventListener("contextmenu",event=>event.preventDefault());
window.addEventListener("resize",()=>closeCellMenu());
$("formula").addEventListener("keydown",e=>{
 if(e.key==="Enter"||e.key==="Escape"){
  e.preventDefault();const [r,c]=selected;
  if(e.key==="Escape")$("formula").value=cells[r][c];
  $("formula").blur();focusGridCell(r,c);
 }
});
$("formula").addEventListener("change",()=>{const [r,c]=selected;cells[r][c]=$("formula").value;renderGrid()});
$("challenge").onchange=e=>{practiceIndex=Number(e.target.value);setupTask(tasks[practiceIndex])};
$("creativeProject").onchange=event=>{creativeIndex=Number(event.target.value);switchMode("creative")};
document.querySelectorAll(".mode").forEach(btn=>btn.onclick=()=>{if(btn.dataset.mode!==mode)switchMode(btn.dataset.mode)});
$("help").onclick=()=>$("guide").showModal();
const soundPreference="formelraster-sound-v1";
let soundEnabled=false,audioContext;
try{soundEnabled=localStorage.getItem(soundPreference)==="on"}catch{}
const reducedMotion=matchMedia("(prefers-reduced-motion: reduce)");
const motionPreference="formelraster-motion-v1";
const motionPicker=$("motionPicker");
try{const saved=localStorage.getItem(motionPreference);if(["system","on","off"].includes(saved))motionPicker.value=saved}catch{}
function updateMotion(){document.documentElement.dataset.motion=motionPicker.value}
motionPicker.onchange=()=>{
 updateMotion();try{localStorage.setItem(motionPreference,motionPicker.value)}catch{}
 if(motionPicker.value==="off")$("sheetEffects").removeAttribute("data-effect");
};
updateMotion();
function refreshSound(){ $("soundToggle").setAttribute("aria-pressed",String(soundEnabled)); }
function playCue(kind){
 if(!soundEnabled)return;
 const Audio=window.AudioContext||window.webkitAudioContext;if(!Audio)return;
 try{
  audioContext??=new Audio();
  if(audioContext.state==="suspended")audioContext.resume().catch(()=>{});
  const notes=kind==="success"?[523,659,784]:kind==="retry"?[290,230]:kind==="task"?[392,523]:[580];
  const time=audioContext.currentTime;
  notes.forEach((hz,i)=>{
   const oscillator=audioContext.createOscillator(),volume=audioContext.createGain();
   const start=time+i*.062;
   oscillator.type="sine";oscillator.frequency.value=hz;
   volume.gain.setValueAtTime(.001,start);volume.gain.exponentialRampToValueAtTime(.075,start+.013);
   volume.gain.exponentialRampToValueAtTime(.001,start+.15);
   oscillator.connect(volume);volume.connect(audioContext.destination);oscillator.start(start);oscillator.stop(start+.16);
  });
 }catch{}
}
const motionClass={task:"cue-task",success:"cue-success",retry:"cue-retry",cell:"cue-cell"};
function signal(kind,target){
 if(motionPicker.value==="on"||(motionPicker.value==="system"&&!reducedMotion.matches)){
  const effect=motionClass[kind];target.classList.remove(effect);void target.offsetWidth;target.classList.add(effect);
  target.addEventListener("animationend",()=>target.classList.remove(effect),{once:true});
  const overlay=$("sheetEffects");
  overlay.removeAttribute("data-effect");void overlay.offsetWidth;overlay.dataset.effect=kind;
 }
 playCue(kind);
}
$("soundToggle").onclick=()=>{
 soundEnabled=!soundEnabled;refreshSound();try{localStorage.setItem(soundPreference,soundEnabled?"on":"off")}catch{}
 if(soundEnabled)playCue("cell");
};
refreshSound();
$("export").onclick=()=>{
 const rows=cells.map((row,r)=>row.map((_,c)=>{try{return String(cellValue(r,c))}catch(error){return error.message}}));
 const csv=rows.map(row=>row.map(v=>'"'+v.replaceAll('"','""')+'"').join(";")).join("\r\n");
 const url=URL.createObjectURL(new Blob(["\uFEFF",csv],{type:"text/csv;charset=utf-8"}));
 const a=document.createElement("a");a.href=url;a.download="formelraster-tabelle.csv";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 message("CSV mit den berechneten Zellwerten gespeichert.","good");
};
const playableFunctions=[
 ["SUMME","SUM","Mathematik","Addiert Zahlen und Zellbereiche.","=SUMME(12;8;5)","25","SUMME(B2:B5)"],
 ["MITTELWERT","AVERAGE","Statistik","Berechnet den Durchschnitt der Zahlen.","=MITTELWERT(10;20;30)","20","MITTELWERT(B2:B6)"],
 ["MIN","MIN","Statistik","Findet den kleinsten Zahlenwert.","=MIN(18;9;24)","9","MIN(B2:B6)"],
 ["MAX","MAX","Statistik","Findet den größten Zahlenwert.","=MAX(18;9;24)","24","MAX(B2:B6)"],
 ["ANZAHL","COUNT","Statistik","Zählt die Zahlen in einer Liste oder in einem Zellbereich.","=ANZAHL(42;8;\"Text\")","2","ANZAHL(B2:B6)"],
 ["WENN","IF","Logik","Gibt je nach Bedingung eines von zwei Ergebnissen zurück.","=WENN(8>5;\"Bestanden\";\"Offen\")","Bestanden","WENN(B2>=10;\"OK\";\"Prüfen\")"],
 ["RUNDEN","ROUND","Mathematik","Rundet eine Zahl auf die gewählte Anzahl Dezimalstellen.","=RUNDEN(3,146;2)","3,15","RUNDEN(B2;2)"],
 ["ZÄHLENWENN","COUNTIF","Statistik","Zählt Werte, die dem Suchwert entsprechen.","=ZÄHLENWENN(\"Retour\";\"Retour\")","1","ZÄHLENWENN(B2:B7;\"Retour\")"],
 ["SUMMEWENN","SUMIF","Mathematik","Addiert nur Werte aus Zeilen, die zum Suchwert passen.","=SUMMEWENN(\"Online\";\"Online\";85)","85","SUMMEWENN(A2:A7;\"Online\";B2:B7)"],
 ["XVERWEIS","XLOOKUP","Nachschlagen","Sucht nach einer exakten Übereinstimmung und liefert den dazugehörigen Wert.","=XVERWEIS(\"P-42\";\"P-42\";18,5)","18,5","XVERWEIS(\"P-42\";A2:A6;B2:B6)"]
];
const categoryNames={"Math and trigonometry":"Mathematik & Trigonometrie","Statistical":"Statistik","Lookup and reference":"Nachschlagen & Verweisen","Financial":"Finanzen","Engineering":"Technik","Logical":"Logik","Information":"Information","Date and time":"Datum & Uhrzeit","Text":"Text","Compatibility":"Kompatibilität","Database":"Datenbank","Cube":"Cube","Web":"Web"};
const playableByEnglish=new Map(playableFunctions.map(entry=>[entry[1],entry]));
let glossaryScope="game";
const searchable=value=>normalized(value).normalize("NFD").replace(/\p{Diacritic}/gu,"").replaceAll("ß","ss");
function glossNode(tag,value,className){
 const node=document.createElement(tag);node.textContent=value;if(className)node.className=className;return node;
}
function tryFormula(formula){
 if($("formula").disabled)return;
 let [r,c]=selected;
 if(cells[r][c]!==""){
  const candidates=cells.flatMap((row,rr)=>row.map((value,cc)=>({value,rr,cc})));
  const empty=candidates.find(cell=>cell.value==="");
  if(!empty){$("glossaryNote").textContent="Die Tabelle ist voll. Leere zuerst eine Zelle.";return}
  r=empty.rr;c=empty.cc;
 }
 cells[r][c]=formula;renderGrid();selectCell(r,c);$("glossary").close();selectCell(r,c);
 message("Beispielformel in "+nameOf(r,c)+" eingefügt.","good");
}
function drawGlossary(){
 const words=searchable($("functionSearch").value).split(/\s+/).filter(Boolean), list=$("functionResults"), isGame=glossaryScope==="game";
 const records=isGame?playableFunctions.map(p=>[p[0],p[1],p[2],null,0]):excelCatalog;
 const entries=records.filter(([name,en,category])=>{
  const play=playableByEnglish.get(en);
  const haystack=searchable([name,en,categoryNames[category]??category,play?.[3]??""].join(" "));
  return words.every(word=>haystack.includes(word));
 }).sort((a,b)=>a[0].localeCompare(b[0],"de"));
 list.replaceChildren();$("functionCount").textContent=entries.length+(entries.length===1?" Funktion":" Funktionen");
 if(!entries.length){list.append(glossNode("p","Keine Treffer. Versuch es mit einem anderen Funktionsnamen."));return}
 const fragment=document.createDocumentFragment();
 for(const [name,en,category,path,german] of entries){
  const playable=playableByEnglish.get(en),card=glossNode("article","","glossary-card"),top=glossNode("div","","glossary-row");
  top.append(glossNode("h3",name+(en&&en!==name?" · "+en:"")),glossNode("span",playable?"Im Spiel":categoryNames[category]??category,"function-badge"));
  card.append(top);
  if(playable){
   card.append(glossNode("p",playable[3]));
   const example=glossNode("div","","glossary-row");
   example.append(glossNode("code",playable[4]+" → "+playable[5]));
   const insert=glossNode("button","Ausprobieren");insert.type="button";insert.disabled=$("formula").disabled;
   insert.addEventListener("click",()=>tryFormula(playable[4]));example.append(insert);card.append(example);
   card.append(glossNode("p","Mit Zellen: ="+playable[6]));
  }
  if(path){
   const link=glossNode("a","Erklärung bei Microsoft ↗");
   link.href="https://support.microsoft.com/"+(german?"de-de":"en-us")+"/excel/"+path;
   link.target="_blank";link.rel="noopener noreferrer";card.append(link);
  }
  fragment.append(card);
 }
 list.append(fragment);
}
$("glossaryOpen").onclick=()=>{
 $("glossaryNote").textContent=mode==="ranked"&&phase==="playing"?"Die Ranked-Uhr läuft beim Nachschlagen weiter.":$("formula").disabled?"Starte eine Runde oder öffne Training, um Beispiele einzusetzen.":"Beispiele werden in eine freie Zelle eingesetzt.";
 drawGlossary();$("glossary").showModal();$("functionSearch").focus();
};
$("functionSearch").addEventListener("input",()=>{drawGlossary();$("glossaryBody").scrollTop=0});
document.querySelectorAll("[data-scope]").forEach(button=>button.onclick=()=>{
 glossaryScope=button.dataset.scope;
 document.querySelectorAll("[data-scope]").forEach(item=>item.setAttribute("aria-pressed",String(item===button)));
 drawGlossary();$("glossaryBody").scrollTop=0;
});
tasks.forEach((t,i)=>$("challenge").add(new Option(t.tier+" · "+t.title,i)));
creativeProjects.forEach((project,i)=>$("creativeProject").add(new Option(project.title,i)));
makeGrid();switchMode("ranked");
