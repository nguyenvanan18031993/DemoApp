/* Report generation stays on this device. Dependencies: local JSZip and browser XML APIs. */
(() => {
  'use strict';
  const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const P='http://schemas.openxmlformats.org/package/2006/relationships';
  const CT='http://schemas.openxmlformats.org/package/2006/content-types';
  const STATUS=['Đã giải ngân','Chưa giải ngân','Đã tất toán'];
  const xml = source => {
    if(/<!DOCTYPE|<!ENTITY/i.test(source)) throw Error('Tệp XML chứa cấu trúc không được hỗ trợ.');
    const doc=new DOMParser().parseFromString(source,'application/xml');
    if(doc.getElementsByTagName('parsererror').length) throw Error('Không đọc được cấu trúc XML trong tệp.');
    return doc;
  };
  const serialize = doc => new XMLSerializer().serializeToString(doc);
  const kids=(el,name,ns=W)=>Array.from(el?.childNodes??[]).filter(n=>n.nodeType===1&&n.localName===name&&n.namespaceURI===ns);
  const one=(el,name,ns=W)=>kids(el,name,ns)[0];
  const all=(el,name,ns=W)=>Array.from(el.getElementsByTagNameNS(ns,name));
  const text=el=>all(el,'t').map(t=>t.textContent).join('');
  const attr=(el,key)=>el?.getAttributeNS(W,key);
  const wa=(el,key,value)=>el.setAttributeNS(W,'w:'+key,String(value));
  const make=(doc,name,ns=W)=>doc.createElementNS(ns,(ns===W?'w:':'')+name);
  function prop(parent,name,value){
    let el=one(parent,name);
    if(!el){el=make(parent.ownerDocument,name);parent.appendChild(el);}
    if(value!==undefined)wa(el,'val',value);
    return el;
  }
  function ppr(p){
    let x=one(p,'pPr');
    if(!x){x=make(p.ownerDocument,'pPr');p.insertBefore(x,p.firstChild);}
    return x;
  }
  function runstyle(p){return (one(one(p,'r'),'rPr')||one(one(p,'pPr'),'rPr'))?.cloneNode(true)||make(p.ownerDocument,'rPr');}
  function runFor(p,value,style){
    const r=make(p.ownerDocument,'r'),rp=style.cloneNode(true);
    const color=prop(rp,'color','000000');
    for(const a of Array.from(color.attributes))if(a.localName!=='val')color.removeAttributeNode(a);
    prop(rp,'sz',20);prop(rp,'szCs',20);r.appendChild(rp);
    String(value??'').split('\n').forEach((line,i)=>{
      if(i)r.appendChild(make(p.ownerDocument,'br'));
      const t=make(p.ownerDocument,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=line;r.appendChild(t);
    });
    return r;
  }
  function setp(p,value){
    const style=runstyle(p),notes=kids(p,'r').filter(r=>all(r,'footnoteReference').length).map(r=>r.cloneNode(true));
    for(const c of Array.from(p.childNodes))if(!['pPr','bookmarkStart','bookmarkEnd'].includes(c.localName))p.removeChild(c);
    p.appendChild(runFor(p,value,style));notes.forEach(r=>p.appendChild(r));
    return p;
  }
  function setcell(cell,value){
    const paras=kids(cell,'p'),p=paras[0]||cell.appendChild(make(cell.ownerDocument,'p'));
    setp(p,value);paras.slice(1).forEach(p=>p.remove());return p;
  }
  const rows=t=>kids(t,'tr'),cells=r=>kids(r,'tc');
  function newrow(proto,values){
    const r=proto.cloneNode(true);cells(r).forEach((c,i)=>setcell(c,values[i]??''));return r;
  }
  function replaceRows(table,old,values,proto){
    values.forEach(v=>table.insertBefore(newrow(proto,v),old[0]));old.forEach(r=>r.remove());
  }
  function keep(row,header=false){
    let pr=one(row,'trPr');if(!pr){pr=make(row.ownerDocument,'trPr');row.insertBefore(pr,row.firstChild);}
    prop(pr,'cantSplit');if(header)prop(pr,'tblHeader');
  }
  function checks(cell,selected){
    const options=kids(cell,'p').filter(p=>text(p).includes('☐'));
    if(options.length!==3)throw Error('Mẫu Word không khớp cấu trúc ô kết luận.');
    options.forEach((p,i)=>{const t=all(p,'t').find(t=>t.textContent.includes('☐'));t.textContent=t.textContent.replace('☐',i===selected?'☒':'☐');});
  }
  const blank=v=>v===undefined||v===null||v==='';
  const money=v=>blank(v)?'':typeof v==='number'?Math.round(v).toLocaleString('vi-VN'):String(v);
  const percent=v=>blank(v)?'':typeof v==='number'?(v*100).toLocaleString('vi-VN',{minimumFractionDigits:1,maximumFractionDigits:1})+'%':String(v);
  const plain=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[đĐ]/g,m=>m==='đ'?'d':'D');
  const safeName=v=>plain(v).replace(/[^A-Za-z0-9_-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,100)||'KH';
  function date(value,date1904=false){
    if(blank(value))return '';
    let d,m,y;
    if(typeof value==='number'){
      const dt=new Date(Date.UTC(date1904?1904:1899,date1904?0:11,date1904?1:30)+Math.floor(value)*86400000);
      d=dt.getUTCDate();m=dt.getUTCMonth()+1;y=dt.getUTCFullYear();
    }else{
      let match=String(value).trim().match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
      if(match)[,d,m,y]=match.map(Number);
      else {match=String(value).trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(match)[,y,m,d]=match.map(Number);}
    }
    const check=new Date(Date.UTC(y,m-1,d));
    if(!y||check.getUTCFullYear()!==y||check.getUTCMonth()+1!==m||check.getUTCDate()!==d)throw Error(`Ngày không hợp lệ: ${value}`);
    return `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}/${y}`;
  }
  function resolvePart(base,target){
    const result=[];
    for(const bit of (target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target).split('/')){
      if(bit==='..')result.pop();else if(bit&&bit!=='.')result.push(bit);
    }
    return result.join('/');
  }
  async function readXml(zip,path){const file=zip.file(path);if(!file)throw Error(`Thiếu thành phần ${path}.`);return xml(await file.async('string'));}
  async function readWorkbook(bytes){
    let zip;try{zip=await JSZip.loadAsync(bytes);}catch{throw Error('Không đọc được tệp Excel. Hãy chọn tệp .xlsx không đặt mật khẩu.');}
    const workbook=await readXml(zip,'xl/workbook.xml'),rels=await readXml(zip,'xl/_rels/workbook.xml.rels');
    const relationships=new Map(all(rels,'Relationship',P).map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));
    const styles=zip.file('xl/styles.xml')?await readXml(zip,'xl/styles.xml'):null;
    const formats=new Map(styles?all(styles,'numFmt',S).map(f=>[Number(f.getAttribute('numFmtId')),f.getAttribute('formatCode')]):[]);
    const xfs=styles?kids(one(styles.documentElement,'cellXfs',S),'xf',S):[];
    const strings=zip.file('xl/sharedStrings.xml')?all(await readXml(zip,'xl/sharedStrings.xml'),'si',S).map(si=>all(si,'t',S).map(t=>t.textContent).join('')):[];
    const date1904=['1','true'].includes(one(workbook.documentElement,'workbookPr',S)?.getAttribute('date1904'));
    const sheets=[];
    for(const entry of kids(one(workbook.documentElement,'sheets',S),'sheet',S)){
      const target=relationships.get(entry.getAttributeNS(R,'id'));
      if(!target)throw Error(`Thiếu liên kết sheet ${entry.getAttribute('name')}.`);
      const path=resolvePart('xl/workbook.xml',target),doc=await readXml(zip,path),values={},errors=[],formulas={};
      for(const cell of all(doc,'c',S)){
        const a=cell.getAttribute('r'),type=cell.getAttribute('t');let v=one(cell,'v',S)?.textContent;
        const formula=one(cell,'f',S);if(formula)formulas[a]=formula.textContent;
        if(type==='s')v=strings[Number(v)]??'';
        else if(type==='inlineStr')v=all(cell,'t',S).map(t=>t.textContent).join('');
        else if(type==='e'){errors.push({cell:a,value:v});}
        else if(v!==undefined&&v!==''&&!['str','b','d'].includes(type)){
          const n=Number(v);if(Number.isFinite(n))v=n;
          const style=xfs[Number(cell.getAttribute('s')||0)],format=formats.get(Number(style?.getAttribute('numFmtId')));
          if(['C13','C14','C16'].includes(a)&&typeof v==='number'&&/^0+$/.test(format??''))v=String(v).padStart(format.length,'0');
        }
        if(v!==undefined)values[a]=v;
      }
      // Recreate the known template's derived fields when Excel omitted cached results.
      for(let r=47;r<=51;r++)if(blank(values[`F${r}`])&&formulas[`F${r}`]&&typeof values[`D${r}`]==='number'&&typeof values[`E${r}`]==='number'&&values[`D${r}`]!==0)values[`F${r}`]=values[`E${r}`]/values[`D${r}`];
      for(const col of ['D','E'])if(blank(values[col+'52'])&&formulas[col+'52'])values[col+'52']=[47,48,49,50,51].reduce((s,r)=>s+(typeof values[col+r]==='number'?values[col+r]:0),0);
      sheets.push({name:entry.getAttribute('name'),entry,path,values,errors,formulas,doc});
    }
    const customers=sheets.filter(s=>!['HUONG_DAN','TONG_HOP'].includes(s.name.toUpperCase())&&plain(s.values.B12??'').toLowerCase().includes('ten khach hang')&&!blank(s.values.C12));
    return {bytes,zip,workbook,rels,sheets,customers,date1904};
  }
  function validateCustomer(c,options={},date1904=false){
    const errors=[];const v=a=>c.values[a]??'';
    for(const error of c.errors)if(!['C91','C98'].includes(error.cell))errors.push(`Ô ${error.cell} có lỗi ${error.value}. Sửa trong Excel và tải lại.`);
    const status=options.status??v('C29');
    if(!blank(status)&&!STATUS.includes(status))errors.push('Chọn lại tình trạng khoản vay.');
    for(const [address,key] of [['C90','conclusion1'],['C97','conclusion2']]){
      const code=options[key]??v(address);if(!blank(code)&&!/^KL[123]$/.test(code))errors.push(`Chọn lại kết luận tại ${address}.`);
    }
    for(const a of ['C9','C23','C28','C94','C101',...Array.from({length:5},(_,i)=>'G'+(47+i))]){
      if(!blank(v(a)))try{date(v(a),date1904);}catch{errors.push(`Ô ${a}: ngày không hợp lệ (${v(a)}).`);}
    }
    const notes=[55,56,57,58,59].some(r=>/^\s*(?:lược bỏ|bỏ ô này|xóa ô này)(?:\s|[,.!?]|$)/iu.test(v('C'+r)));
    if(notes&&options.omitNarratives!==true)errors.push('C55–C59 chứa ghi chú yêu cầu lược bỏ. Xác nhận bỏ các đoạn này hoặc sửa nội dung trong Excel.');
    return errors;
  }
  function findings(c){
    const result=[],v=a=>String(c.values[a]??'').trim();
    for(const [lo,hi] of [[63,68],[72,77],[81,86]])for(let r=lo;r<=hi;r++){
      if(v('D'+r).toLocaleLowerCase('vi')!=='không tuân thủ')continue;
      const topic=v('B'+r),detail=v('C'+r);
      const report=/biên bản kiểm tra sau cho vay/iu.test(topic+' '+detail);
      const key=report?'bien-ban-'+(detail.match(/\d{1,2}\/\d{4}/u)?.[0]??''):topic+'\n'+detail;
      if(result.some(f=>f.key===key))continue;
      result.push({key,text:(report?'Biên bản kiểm tra sau cho vay':topic.replace(/[.:;\s]+$/u,''))+': '+(detail||'Chưa có chi tiết trong hồ sơ.')});
    }
    return result;
  }
  async function openTemplate(template){
    if(!template?.data)throw Error('Hãy chọn mẫu Word trước khi xuất báo cáo.');
    let zip;try{zip=await JSZip.loadAsync(template.data,{base64:!!template.base64});}catch{throw Error('Không đọc được mẫu Word. Hãy chọn tệp .docx không đặt mật khẩu.');}
    const doc=await readXml(zip,'word/document.xml'),body=one(doc.documentElement,'body'),ps=kids(body,'p'),ts=kids(body,'tbl');
    if(ts.length!==7||ps.length!==24||all(body,'sdt').length!==9||!["Dư nợ vay thế chấp","HMTC tiêu dùng có TSBĐ","HM Thẻ tín dụng của KH"].every(label=>text(ts[2]).includes(label)))throw Error('Mẫu Word không đúng bố cục đã mapping. Hãy chọn bản mẫu Word mới nhất đã dùng cho báo cáo.');
    return {zip,doc,body,ps,ts};
  }
  async function generateWord(customer,options={},date1904=false,template){
    const issues=validateCustomer(customer,options,date1904);if(issues.length)throw Error(`${customer.name}: ${issues.join(' ')}`);
    const {zip,doc,ps,ts}=await openTemplate(template);
    const d={...customer.values,C29:options.status??customer.values.C29,C90:options.conclusion1??customer.values.C90,C97:options.conclusion2??customer.values.C97};
    const v=a=>d[a]??'',dt=v=>date(v,date1904);
    const direct=kids(ps[0],'r').flatMap(r=>kids(r,'t'));
    let department=String(v('C7'));
    if(/^\d+$/.test(department.trim()))department=direct[0].textContent.split(/[.…]/)[0].trim()+' '+department.trim();
    direct.forEach((t,i)=>t.textContent=i===0?department:'');
    const meta=ps[1],r0=runstyle(meta),r1=one(kids(meta,'r').at(-1),'rPr').cloneNode(true);
    setp(meta,'');kids(meta,'r').forEach(r=>r.remove());
    meta.appendChild(runFor(meta,'Tên cán bộ rà soát: '+v('C8'),r0));
    const tr=make(doc,'r');tr.appendChild(make(doc,'tab'));meta.appendChild(tr);
    const parts=dt(v('C9')).split('/');
    meta.appendChild(runFor(meta,parts.length===3?`Ngày ${parts[0]} tháng ${parts[1]} năm ${parts[2]}`:'Ngày        tháng        năm        ',r1));
    const tabs=prop(ppr(meta),'tabs'),tab=make(doc,'tab');wa(tab,'val','right');wa(tab,'pos','9072');tabs.appendChild(tab);
    let rr=rows(ts[0]),cc=rr.map(cells);
    for(const [cell,value] of [[cc[0][1],v('C12')],[cc[0][2],'CCCD:\n'+v('C13')],[cc[0][4],v('C14')],[cc[1][1],v('C15')],[cc[1][2],'CCCD:\n'+v('C16')],[cc[2][1],v('C17')],[cc[2][3],v('C18')]])setcell(cell,value);
    rr=rows(ts[1]);setcell(cells(rr[0])[0],`Căn cứ phê duyệt tại ${v('C21')} số ${v('C22')} ngày ${dt(v('C23'))}.`);
    for(const [index,a] of [[1,'C24'],[2,'C25'],[3,'C26'],[4,'C27'],[5,'C28'],[7,'C30'],[9,'C33']])setcell(cells(rr[index])[1],a==='C24'?(blank(v(a))?'':money(v(a))+' đồng'):a==='C28'?dt(v(a)):v(a));
    setcell(cells(rr[8])[1],[v('C31'),v('C32')].filter(x=>!blank(x)).join(' – '));
    checks(cells(rr[6])[1],STATUS.indexOf(v('C29')));
    kids(cells(rows(ts[2])[0])[1],'p').forEach((p,i)=>setp(p,['Dư nợ vay thế chấp','HMTC tiêu dùng có TSBĐ','HM Thẻ tín dụng của KH'][i]+': '+(blank(v('C'+(36+i)))?'':money(v('C'+(36+i)))+' đồng')));
    let t=ts[3];rr=rows(t);const assets=[];
    for(let r=47;r<=51;r++)if(!blank(v('B'+r)))assets.push([assets.length+1,v('B'+r),v('C'+r),money(v('D'+r)),money(v('E'+r)),percent(v('F'+r)),dt(v('G'+r))]);
    replaceRows(t,rr.slice(1,3),assets,rr[1]);setcell(cells(rr[3])[1],money(v('D52')));setcell(cells(rr[3])[2],money(v('E52')));
    setcell(cells(rr[0])[3],'Giá trị TS\n(đồng)');setcell(cells(rr[0])[4],'Dư nợ\n(đồng)');
    const weights=[.42,.58,1.25,1.13,1.13,.55,1.15],grid=kids(one(t,'tblGrid'),'gridCol'),total=grid.reduce((n,g)=>n+Number(attr(g,'w')),0),sum=weights.reduce((a,b)=>a+b,0),widths=weights.map(w=>Math.round(w/sum*total));
    widths[6]=total-widths.slice(0,6).reduce((a,b)=>a+b,0);grid.forEach((g,i)=>wa(g,'w',widths[i]));wa(prop(one(t,'tblPr'),'tblLayout'),'type','fixed');
    for(const r of rows(t)){
      let column=0;
      for(const c of cells(r)){
        const cp=one(c,'tcPr'),span=Number(attr(one(cp,'gridSpan'),'val')||1),cw=prop(cp,'tcW');wa(cw,'type','dxa');wa(cw,'w',widths.slice(column,column+span).reduce((a,b)=>a+b,0));
        kids(c,'p').forEach(p=>{const pp=ppr(p),ind=prop(pp,'ind');['left','right','firstLine','hanging'].forEach(a=>wa(ind,a,0));prop(pp,'jc',column===2&&r!==rows(t)[0]?'left':'center');});column+=span;
      }
    }
    if(options.omitNarratives)ps.slice(16,19).forEach(p=>p.remove());
    else {setp(ps[16],'Căn cứ hồ sơ: '+v('C55'));setp(ps[17],`ĐVKD không cung cấp các hồ sơ: ${v('C56')}\nHồ sơ không bổ sung được: ${v('C57')}\nLý do: ${v('C58')}`);setp(ps[18],'Phòng TĐTD có ý kiến đánh giá tuân thủ như sau: '+v('C59'));}
    t=ts[4];rr=rows(t);
    for(const [lo,hi,old] of [[63,68,rr.slice(2,5)],[72,77,rr.slice(6,9)]]){
      const values=[];for(let r=lo;r<=hi;r++)if(!blank(v('B'+r)))values.push([values.length+1,v('B'+r),v('C'+r),v('D'+r)]);replaceRows(t,old,values,old[0]);
    }
    t=ts[5];rr=rows(t);const monitoring=[];
    for(let r=81;r<=86;r++)if(!blank(v('B'+r))&&(!blank(v('C'+r))||!blank(v('D'+r))))monitoring.push([monitoring.length+1,v('B'+r),v('C'+r),v('D'+r)]);
    replaceRows(t,rr.slice(1),monitoring,rr[1]);
    const first=cells(rows(t)[0])[0];kids(first,'p').forEach(p=>{const ind=prop(ppr(p),'ind');['left','right','firstLine','hanging'].forEach(a=>wa(ind,a,0));});
    const margins=prop(one(first,'tcPr'),'tcMar');for(const side of ['left','right']){const m=prop(margins,side);wa(m,'w',30);wa(m,'type','dxa');}
    for(const [row,code,proposal,name,day] of [[1,'C90','C92','C93','C94'],[3,'C97','C99','C100','C101']]){
      const cell=cells(rows(ts[6])[row])[0];checks(cell,blank(v(code))?-1:Number(v(code).slice(-1))-1);
      const paras=kids(cell,'p'),sign=paras.find(p=>text(p)==='Cấp kiểm soát TĐTD KHCN');
      if(!blank(v(proposal))){const p=paras.find(p=>text(p).includes('ĐVKD tuân thủ đầy đủ')).cloneNode(true);setp(p,'Đề xuất, kiến nghị: '+v(proposal));cell.insertBefore(p,sign);}
      paras.forEach(p=>{if(text(p).startsWith('Họ và tên:'))setp(p,'Họ và tên: '+v(name));else if(text(p).startsWith('Ngày:'))setp(p,'Ngày: '+dt(v(day)));});
    }
    ts.forEach((t,i)=>rows(t).forEach((r,j)=>keep(r,j===0&&i>=3&&i<=5)));
    ps.slice(-2).forEach(p=>{if(!text(p)){const pp=ppr(p),sp=prop(pp,'spacing');wa(sp,'before',0);wa(sp,'after',0);wa(sp,'line',20);wa(sp,'lineRule','exact');const rp=prop(pp,'rPr');prop(rp,'sz',2);prop(rp,'szCs',2);}});
    [6,7,8,10,12,14,15,19,21].forEach(i=>prop(ppr(ps[i]),'keepNext'));
    [12,21].forEach(i=>prop(ppr(ps[i]),'pageBreakBefore'));
    for(const [table,indices] of [[ts[4],[1]],[ts[6],[0,2]]])indices.forEach(i=>cells(rows(table)[i]).forEach(c=>kids(c,'p').forEach(p=>prop(ppr(p),'keepNext'))));
    zip.file('word/document.xml',serialize(doc));
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  }
  const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  async function summaryWorkbook(source,customers){
    const zip=await JSZip.loadAsync(source.bytes),wb=await readXml(zip,'xl/workbook.xml'),rels=await readXml(zip,'xl/_rels/workbook.xml.rels'),ct=await readXml(zip,'[Content_Types].xml');
    const styles=await readXml(zip,'xl/styles.xml'),root=styles.documentElement;
    const fonts=one(root,'fonts',S),fills=one(root,'fills',S),xfs=one(root,'cellXfs',S);
    function append(collection,source){const id=kids(collection,collection.localName==='fonts'?'font':collection.localName==='fills'?'fill':'xf',S).length;collection.appendChild(collection.ownerDocument.importNode(xml(source).documentElement,true));collection.setAttribute('count',id+1);return id;}
    const font=append(fonts,`<font xmlns="${S}"><sz val="11"/><color rgb="FF233249"/><name val="Arial"/></font>`);
    const titleFont=append(fonts,`<font xmlns="${S}"><b/><sz val="15"/><color rgb="FF233249"/><name val="Arial"/></font>`);
    const headerFont=append(fonts,`<font xmlns="${S}"><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>`);
    const blue=append(fills,`<fill xmlns="${S}"><patternFill patternType="solid"><fgColor rgb="FF24476B"/><bgColor indexed="64"/></patternFill></fill>`);
    const stripe=append(fills,`<fill xmlns="${S}"><patternFill patternType="solid"><fgColor rgb="FFF2F4F7"/><bgColor indexed="64"/></patternFill></fill>`);
    const style=(fontId,fillId,align='left',wrap=true)=>append(xfs,`<xf xmlns="${S}" numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="${align}" vertical="top" wrapText="${wrap?1:0}"/></xf>`);
    const normal=style(font,0),alternate=style(font,stripe),heading=style(headerFont,blue,'center'),title=style(titleFont,0,'left',false);
    const sheets=one(wb.documentElement,'sheets',S),existing=kids(sheets,'sheet',S),oldOrder=existing.map(s=>s.getAttribute('name'));
    let summary=existing.find(s=>s.getAttribute('name').toUpperCase()==='TONG_HOP'),path;
    if(summary){const rel=all(rels,'Relationship',P).find(r=>r.getAttribute('Id')===summary.getAttributeNS(R,'id'));path=resolvePart('xl/workbook.xml',rel.getAttribute('Target'));sheets.appendChild(summary);summary.setAttribute('name','TONG_HOP');summary.removeAttribute('state');}
    else{
      let index=1;while(zip.file(`xl/worksheets/sheet${index}.xml`))index++;path=`xl/worksheets/sheet${index}.xml`;
      let id=1;const used=new Set(all(rels,'Relationship',P).map(r=>r.getAttribute('Id')));while(used.has('rId'+id))id++;
      const relationship=rels.createElementNS(P,'Relationship');relationship.setAttribute('Id','rId'+id);relationship.setAttribute('Type',R+'/worksheet');relationship.setAttribute('Target',`worksheets/sheet${index}.xml`);rels.documentElement.appendChild(relationship);
      summary=wb.createElementNS(S,'sheet');summary.setAttribute('name','TONG_HOP');summary.setAttribute('sheetId',Math.max(0,...existing.map(s=>Number(s.getAttribute('sheetId'))))+1);summary.setAttributeNS(R,'r:id','rId'+id);sheets.appendChild(summary);
      const content=ct.createElementNS(CT,'Override');content.setAttribute('PartName','/'+path);content.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');ct.documentElement.appendChild(content);
    }
    const order=kids(sheets,'sheet',S).map(s=>s.getAttribute('name'));
    const newIndex=old=>order.findIndex(name=>name.toUpperCase()===oldOrder[old]?.toUpperCase());
    for(const name of all(wb,'definedName',S))if(name.hasAttribute('localSheetId')){const index=newIndex(Number(name.getAttribute('localSheetId')));if(index>=0)name.setAttribute('localSheetId',index);}
    for(const view of all(wb,'workbookView',S))for(const key of ['activeTab','firstSheet'])if(view.hasAttribute(key)){const index=newIndex(Number(view.getAttribute(key)));if(index>=0)view.setAttribute(key,index);}
    const cell=(address,value,s)=>`<c r="${address}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
    const rows=[`<row r="2" ht="30" customHeight="1">${cell('B2','TỔNG HỢP LỖI HỒ SƠ',title)}</row>`,`<row r="4" ht="28" customHeight="1">${['Tên KH','Tên đơn vị kinh doanh','Lỗi phát sinh của hồ sơ'].map((v,i)=>cell('BCD'[i]+'4',v,heading)).join('')}</row>`];
    const records=customers.map(c=>{const errors=findings(c);return [c.values.C12??'',c.values.C17??'',errors.length?errors.map((f,i)=>`${i+1}. ${f.text}`).join('\n'):'Không ghi nhận lỗi trong các nội dung đã đánh giá.'];});
    records.forEach((record,i)=>{const r=i+5;const lines=Math.max(...record.map((v,j)=>String(v).split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(line.length/[24,32,94][j])),0)));const ht=Math.min(409,Math.max(38,lines*16+10));rows.push(`<row r="${r}" ht="${ht}" customHeight="1">${record.map((v,j)=>cell('BCD'[j]+r,v,i%2?alternate:normal)).join('')}</row>`);});
    const last=records.length+4,pane=records.length>10?'<pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>':'';
    const sheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}"><sheetPr><tabColor rgb="FF24476B"/></sheetPr><dimension ref="B2:D${last}"/><sheetViews><sheetView showGridLines="0" workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="16"/><cols><col min="1" max="1" width="3" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/><col min="3" max="3" width="36" customWidth="1"/><col min="4" max="4" width="104" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><autoFilter ref="B4:D${last}"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`;
    zip.file(path,sheet);zip.file('xl/workbook.xml',serialize(wb));zip.file('xl/_rels/workbook.xml.rels',serialize(rels));zip.file('[Content_Types].xml',serialize(ct));zip.file('xl/styles.xml',serialize(styles));
    // Update the visible sheet-name inventory when present; all input worksheet parts stay untouched.
    if(zip.file('docProps/app.xml')){
      const app=await readXml(zip,'docProps/app.xml'),vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',ep='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
      const titles=all(app,'TitlesOfParts',ep)[0],vec=titles?all(titles,'vector',vt)[0]:null;
      const pairs=all(app,'HeadingPairs',ep)[0];
      if(vec&&pairs){
        const variants=all(pairs,'variant',vt),items=Array.from(vec.children);let offset=0;
        for(let i=0;i+1<variants.length;i+=2){
          const count=all(variants[i+1],'i4',vt)[0],n=Number(count?.textContent);
          if(!Number.isInteger(n)||n<0)break;
          if(all(variants[i],'lpstr',vt)[0]?.textContent==='Worksheets'&&items.length>=offset+n){
            const next=items[offset+n]??null;
            items.slice(offset,offset+n).forEach(item=>item.remove());
            for(const name of order){const s=app.createElementNS(vt,'vt:lpstr');s.textContent=name;vec.insertBefore(s,next);}
            count.textContent=order.length;vec.setAttribute('size',vec.children.length);zip.file('docProps/app.xml',serialize(app));break;
          }
          offset+=n;
        }
      }
    }
    return {bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}}),records};
  }
  globalThis.ReportEngine={readWorkbook,validateCustomer,validateTemplate:async template=>{await openTemplate(template);},generateWord,summaryWorkbook,findings,safeName,date,STATUS};
})();
