"use strict";
(() => {
 const key="formelraster-theme";
 const supported=["system","light","dark"];
 const scheme=matchMedia("(prefers-color-scheme: dark)");
 let choice="system";
 try { const stored=localStorage.getItem(key); if(supported.includes(stored))choice=stored; } catch {}
 const apply=()=>{document.documentElement.dataset.theme=choice==="system"?(scheme.matches?"dark":"light"):choice};
 apply();scheme.addEventListener("change",apply);
 document.addEventListener("DOMContentLoaded",()=>{
  const picker=document.getElementById("themePicker");
  picker.value=choice;
  picker.addEventListener("change",()=>{
   choice=picker.value;apply();
   try{localStorage.setItem(key,choice)}catch{}
  });
 });
})();
