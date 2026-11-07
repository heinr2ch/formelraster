"use strict";
function generateEndlessTask(level,random=Math.random){
 const roll=(low,high)=>low+Math.floor(random()*(high-low+1));
 const count=Math.min(9,4+Math.floor((level-1)/2));
 const scale=Math.min(100000,35+level*9);
 const tier=level<4?"Einstieg":level<8?"Aufbau":level<13?"Profi":"Experte";
 const scenes=["Mondkolonie","Nachtmarkt","Polarstation","Roboterliga","Tiefseelabor","Festivalcrew","Luftschiffwerft","Geisterarchiv"];
 const scene=scenes[(level-1+Math.floor(random()*scenes.length))%scenes.length];
 const task=(title,brief,hint,answer,headers,rows)=>({tier,title:scene+" · "+title,brief,hint,answer,rows:[headers,...rows]});
 const end=count+1;
 if(level<=3){
  const rows=Array.from({length:count},(_,i)=>["P-"+(101+i),roll(7,scale)]);
  const numbers=rows.map(row=>row[1]);
  if(level===1)return task("Wareneingang","Wie viele Stück wurden insgesamt geliefert?",`Addiere B2:B${end} mit SUMME.`,numbers.reduce((a,b)=>a+b,0),["Paket","Stück"],rows);
  if(level===2)return task("Höchster Eingang","Wie viele Stück enthält das größte Paket?",`Nutze MAX(B2:B${end}).`,Math.max(...numbers),["Paket","Stück"],rows);
  return task("Stück pro Paket","Wie viele Stück sind im Durchschnitt in einem Paket? Runde auf zwei Dezimalstellen.",`Nutze RUNDEN(MITTELWERT(B2:B${end});2).`,Math.round(numbers.reduce((a,b)=>a+b,0)*100/count)/100,["Paket","Stück"],rows);
 }
 if(level<=7){
   const groups=["Online","Shop","Partner"],group=groups[roll(0,2)];
  const rows=Array.from({length:count},(_,i)=>[i===0?group:groups[roll(0,2)],roll(10,scale),roll(0,12)]);
  const picked=rows.filter(row=>row[0]===group);
  if(level===4)return task("Kanal zählen",`Wie viele Buchungen gehören zu „${group}“?`,`ZÄHLENWENN(A2:A${end};"${group}") zählt passende Zeilen.`,picked.length,["Kanal","Umsatz","Retouren"],rows);
  if(level===5){
   const cutoff=roll(15,scale-5);
   rows[0][1]=cutoff+1;
   return task("Über der Freigabegrenze",`Wie viele Buchungen haben mehr als ${cutoff} Umsatz?`,
    `Prüfe jede Zeile mit WENN(B2>${cutoff};1;0), dann summiere deine Hilfsspalte.`,
    rows.filter(row=>row[1]>cutoff).length,["Kanal","Umsatz","Retouren"],rows);
  }
  if(level===6)return task("Umsatz nach Kanal",`Wie viel Umsatz kommt aus „${group}“?`,`SUMMEWENN(A2:A${end};"${group}";B2:B${end}) summiert nur passende Zeilen.`,picked.reduce((sum,row)=>sum+row[1],0),["Kanal","Umsatz","Retouren"],rows);
  return task("Netto nach Retouren",`Wie hoch ist der Umsatz von „${group}“ nach Abzug seiner Retouren?`,"Summiere Umsatz und Retouren für denselben Kanal separat, dann ziehe die Retouren ab.",picked.reduce((sum,row)=>sum+row[1]-row[2],0),["Kanal","Umsatz","Retouren"],rows);
 }
 if(level<=12){
  const rows=Array.from({length:count},(_,i)=>["A-"+(201+i),roll(1,9+level),roll(5,scale),roll(0,8)]);
  if(level%2===0){
   const [id,quantity,price]=rows[roll(0,rows.length-1)];
   return task("Nachschlagen und rechnen",`Was kosten alle Stück des Artikels ${id} vor Abzug des Rabatts?`,"Suche Stückzahl und Preis mit zwei XVERWEIS-Formeln und multipliziere beide Ergebnisse.",quantity*price,["Artikel","Stück","Preis/Stück","Rabatt"],rows);
  }
  if(level===11){
   const fee=roll(2,8),limit=roll(4,10);
   return task("Expedition mit Zoll",`Wie hoch ist der Gesamtwert nach Rabatt, inklusive ${fee} Zoll pro Zeile mit mindestens ${limit} Stück?`,
    `Rechne je Zeile Stück × Preis − Rabatt und addiere WENN(B2>=${limit};${fee};0). Summiere die Ergebnisse.`,
    rows.reduce((sum,row)=>sum+row[1]*row[2]-row[3]+(row[1]>=limit?fee:0),0),["Artikel","Stück","Preis/Stück","Rabatt"],rows);
  }
  return task("Der Warenkorb", "Wie hoch ist der Gesamtwert nach Rabatt? Der Rabatt gilt einmal pro Zeile.",
   "Berechne je Zeile Stück × Preis − Rabatt in einer Hilfsspalte und summiere diese.",
   rows.reduce((sum,row)=>sum+row[1]*row[2]-row[3],0),["Artikel","Stück","Preis/Stück","Rabatt"],rows);
 }
 const channels=["Online","Shop","Partner"],channel=channels[roll(0,2)],threshold=roll(4,Math.min(160,7+level));
 const conditionCount=level<17?2:level<23?3:4;
 const minimum=roll(15,scale);
 const rows=Array.from({length:count},(_,i)=>[
  i===0?channel:channels[roll(0,2)],i===0?threshold:roll(1,threshold+12),
  i===0?minimum:roll(10,scale+10),roll(1,Math.floor(scale/4)),i===0?"Ja":roll(0,1)?"Ja":"Nein"
 ]);
 const matches=row=>row[0]===channel&&row[1]>=threshold&&(conditionCount<3||row[4]==="Ja")&&(conditionCount<4||row[2]>=minimum);
 const selectedRows=rows.filter(matches);
 const conditions=`Kanal „${channel}“, mindestens ${threshold} Stück`+(conditionCount>=3?", freigegeben":"")+(conditionCount>=4?` und Stückpreis mindestens ${minimum}`:"");
 const headers=["Kanal","Stück","Preis/Stück","Kosten","Freigabe"];
 const net=row=>row[1]*row[2]-row[3];
 const variant=level<17?0:roll(0,level<23?1:2);
 const adjustment=level>=28?roll(2,17):0;
 const value=row=>net(row)-adjustment*row[1];
 const definition="Netto = Stück × Preis − Kosten"+(adjustment?` − ${adjustment} je Stück.`:".");
 const hint="Prüfe die Bedingungen je Zeile mit verschachteltem WENN. Rechne nur für passende Zeilen und nutze eine Hilfsspalte.";
 if(variant===1)return task("Spitzenwert mit Bedingungen",`Welche passende Zeile hat den höchsten Nettowert? Gib den Wert ein. Kriterien: ${conditions}. ${definition}`,hint,Math.max(...selectedRows.map(value)),headers,rows);
 if(variant===2)return task("Durchschnitt mit Bedingungen",`Wie hoch ist der durchschnittliche Nettowert der passenden Zeilen? Runde auf zwei Dezimalstellen. Kriterien: ${conditions}. ${definition}`,hint,Math.round(selectedRows.reduce((sum,row)=>sum+value(row),0)*100/selectedRows.length)/100,headers,rows);
 return task("Nettowert mit Bedingungen",`Wie hoch ist die Summe der Nettowerte für ${conditions}? ${definition}`,hint,selectedRows.reduce((sum,row)=>sum+value(row),0),headers,rows);
}
