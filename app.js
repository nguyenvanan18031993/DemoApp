(() => {
  'use strict';
  const $=id=>document.getElementById(id),E=globalThis.ReportEngine;
  const state={workbook:null,customers:[],file:null,urls:[],busy:false,templateBusy:false,template:globalThis.REPORT_TEMPLATE_BASE64?{data:globalThis.REPORT_TEMPLATE_BASE64,base64:true,name:'Mẫu Word đã tích hợp'}:null};
  const escape=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const reveal=(id,visible)=>$(id).hidden=!visible;
  function clearDownloads(){state.urls.forEach(URL.revokeObjectURL);state.urls=[];reveal('results',false);$('downloads').replaceChildren();}
  function url(bytes,type){const u=URL.createObjectURL(new Blob([bytes],{type}));state.urls.push(u);return u;}
  function errors(c){return E.validateCustomer(c.source,c.options,state.workbook.date1904);}
  function setSteps(step){[1,2,3].forEach(n=>{$('step'+n).className=n===step?'current':n<step?'complete':'';});}
  function updateActions(){
    const selected=state.customers.filter(c=>c.selected),invalid=selected.filter(c=>errors(c).length);
    $('generate').disabled=state.busy||state.templateBusy||!state.template||!selected.length||!!invalid.length;
    $('export-label').textContent=!state.template?'Chưa chọn mẫu Word':invalid.length?`${invalid.length} hồ sơ cần kiểm tra`:selected.length?`${selected.length} hồ sơ đã chọn`:'Chưa chọn hồ sơ';
    $('export-detail').textContent=!state.template?'Chọn mẫu Word ở phần Tệp nhập liệu để tiếp tục.':invalid.length?'Mở “Kiểm tra” để hoàn tất các mục còn thiếu.':`${selected.length} file Word và 1 Excel có sheet TONG_HOP`;
    $('select-all').textContent=selected.length===state.customers.length?'Bỏ chọn tất cả':'Chọn tất cả';
    reveal('empty-selection',!selected.length);
  }
  function updateTemplate(){
    $('word-name').textContent=state.template?.name||'Mẫu Word';
    $('word-detail').textContent=state.template?'Sẵn sàng dùng cho các báo cáo trong phiên này.':'Chọn mẫu .docx đã mapping một lần cho phiên làm việc.';
    $('template-badge-label').textContent=state.template?(state.template.base64?'Mẫu Word đã tích hợp':'Đã chọn mẫu Word'):'Chưa chọn mẫu Word';
    $('choose-word').textContent=state.template?'Đổi mẫu Word':'Chọn mẫu Word';
  }
  async function loadWord(file){
    if(!file||state.busy||state.templateBusy)return;
    reveal('word-error',false);state.templateBusy=true;$('choose-word').disabled=true;if(state.workbook)updateActions();
    try{
      if(!/\.docx$/i.test(file.name))throw Error('Vui lòng chọn mẫu Word định dạng .docx.');
      if(file.size>20*1024*1024)throw Error('Mẫu Word lớn hơn 20 MB. Hãy chọn bản mẫu đã mapping.');
      const template={data:await file.arrayBuffer(),name:file.name};await E.validateTemplate(template);
      state.template=template;clearDownloads();reveal('generate-error',false);reveal('progress-area',false);if(state.workbook)setSteps(2);updateTemplate();
    }catch(error){$('word-error').textContent=error.message||'Không đọc được mẫu Word.';reveal('word-error',true);}
    finally{state.templateBusy=false;$('choose-word').disabled=false;$('word-input').value='';if(state.workbook)updateActions();}
  }
  function option(value,label,current){return `<option value="${escape(value)}" ${value===current?'selected':''}>${escape(label)}</option>`;}
  function selectOptions(values,current,labels){return (!values.includes(current)&&current!==''?option(current,`Cần chọn lại (${current})`,current):'')+option('','Để trống',current)+values.map((v,i)=>option(v,labels?.[i]??v,current)).join('');}
  function renderCustomers(){
    const focus=document.activeElement?.id;
    $('customer-list').innerHTML=state.customers.map((c,i)=>{
      const issue=errors(c),findings=E.findings(c.source),d=c.source.values,o=c.options;
      return `<article class="customer"><div class="customer-top"><input id="selected-${i}" type="checkbox" data-index="${i}" data-field="selected" ${c.selected?'checked':''} aria-label="Chọn ${escape(d.C12)}"><div class="customer-name"><strong>${escape(d.C12)}</strong><span>${escape(c.source.name)}${d.C14?' · '+escape(d.C14):''}</span></div><div class="customer-unit">${escape(d.C17||'Chưa nhập đơn vị kinh doanh')}</div><div class="customer-actions"><span class="finding-count ${findings.length?'has-findings':''}">${findings.length?findings.length+' lỗi phát sinh':'Không ghi nhận lỗi'}</span><button type="button" class="review-button ${issue.length?'needs-review':''}" data-index="${i}" data-action="review" aria-expanded="${c.open}" aria-controls="review-${i}">${issue.length?'Kiểm tra '+issue.length+' mục':c.open?'Thu gọn':'Xem chi tiết'}</button></div></div><div id="review-${i}" class="customer-details" ${c.open?'':'hidden'}><div class="options-grid"><div class="field"><label for="status-${i}">Tình trạng khoản vay</label><select id="status-${i}" data-index="${i}" data-field="status">${selectOptions(E.STATUS,o.status)}</select></div><div class="field"><label for="conclusion1-${i}">Kết luận cán bộ rà soát</label><select id="conclusion1-${i}" data-index="${i}" data-field="conclusion1">${selectOptions(['KL1','KL2','KL3'],o.conclusion1,['KL1 · Tuân thủ đầy đủ','KL2 · Vi phạm, chưa phải dừng GN','KL3 · Vi phạm, phải dừng GN / chuyển CM'])}</select></div><div class="field"><label for="conclusion2-${i}">Kết luận cấp kiểm soát</label><select id="conclusion2-${i}" data-index="${i}" data-field="conclusion2">${selectOptions(['KL1','KL2','KL3'],o.conclusion2,['KL1 · Tuân thủ đầy đủ','KL2 · Vi phạm, chưa phải dừng GN','KL3 · Vi phạm, phải dừng GN / chuyển CM'])}</select></div></div><label class="omit-control"><input id="omit-${i}" type="checkbox" data-index="${i}" data-field="omitNarratives" ${o.omitNarratives?'checked':''}>Bỏ 5 đoạn căn cứ hồ sơ và ý kiến chung (C55–C59) khỏi Word</label>${issue.length?`<div class="review-errors" role="status">${issue.map(escape).join('<br>')}</div>`:''}${findings.length?`<ul>${findings.map(f=>`<li>${escape(f.text)}</li>`).join('')}</ul>`:'<ul><li>Không ghi nhận lỗi trong các nội dung đã đánh giá.</li></ul>'}</div></article>`;
    }).join('');
    if(focus&&$(focus))$(focus).focus({preventScroll:true});updateActions();
  }
  async function loadFile(file){
    if(state.busy)return;
    reveal('upload-error',false);
    if(!file)return;
    if(!/\.xlsx$/i.test(file.name)){$('upload-error').textContent='Vui lòng chọn tệp .xlsx. Tệp .xls và .xlsm chưa được hỗ trợ.';reveal('upload-error',true);return;}
    if(file.size>30*1024*1024){$('upload-error').textContent='Tệp lớn hơn 30 MB. Hãy tách thành các đợt nhỏ hơn rồi tải lại.';reveal('upload-error',true);return;}
    state.busy=true;$('dropzone').classList.add('loading');$('dropzone').disabled=true;$('change-file').disabled=true;
    try{
      const bytes=await file.arrayBuffer(),wb=await E.readWorkbook(bytes);
      if(!wb.customers.length)throw Error('Không tìm thấy sheet khách hàng có tên tại C12 và nhãn “Tên khách hàng” tại B12. Hãy dùng file Excel mẫu đã mapping.');
      state.workbook=wb;state.file=file;
      state.customers=wb.customers.map(source=>({source,selected:true,open:false,options:{status:String(source.values.C29??'').trim(),conclusion1:String(source.values.C90??'').trim(),conclusion2:String(source.values.C97??'').trim(),omitNarratives:false}}));
      clearDownloads();reveal('generate-error',false);reveal('progress-area',false);
      state.customers.forEach(c=>c.open=errors(c).length>0);
      $('file-name').textContent=file.name;$('file-detail').textContent=`${(file.size/1024).toLocaleString('vi-VN',{maximumFractionDigits:0})} KB · ${wb.customers.length} sheet khách hàng · ${wb.sheets.length-wb.customers.length} sheet khác`;
      $('customer-count').textContent=wb.customers.length;
      reveal('dropzone',false);reveal('file-card',true);reveal('customer-section',true);reveal('intro',false);setSteps(2);
    }catch(error){$('upload-error').textContent=error.message||'Không đọc được tệp Excel.';reveal('upload-error',true);}
    finally{state.busy=false;$('dropzone').classList.remove('loading');$('dropzone').disabled=false;$('change-file').disabled=false;if(state.workbook)renderCustomers();$('file-input').value='';}
  }
  function reset(){
    if(state.busy)return;clearDownloads();state.workbook=null;state.customers=[];state.file=null;
    reveal('dropzone',true);reveal('file-card',false);reveal('customer-section',false);reveal('intro',true);reveal('upload-error',false);$('file-input').value='';setSteps(1);
  }
  function downloadItem(name,label,bytes,kind){
    const href=url(bytes,kind==='xlsx'?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    const link=document.createElement('a');link.className='download-item';link.href=href;link.download=name;
    link.innerHTML=`<span class="download-type ${kind}">${kind.toUpperCase()}</span><span>${escape(label)}</span><span aria-hidden="true">↓</span>`;
    $('downloads').appendChild(link);return href;
  }
  function progress(percent,label){$('progress-fill').style.width=percent+'%';$('progress-text').textContent=label;}
  async function generate(){
    const selected=state.customers.filter(c=>c.selected);if(!selected.length||selected.some(c=>errors(c).length)||state.busy||state.templateBusy||!state.template)return;
    clearDownloads();reveal('generate-error',false);reveal('progress-area',true);state.busy=true;updateActions();
    $('customer-list').inert=true;$('select-all').disabled=true;$('change-file').disabled=true;$('clear-file').disabled=true;$('choose-word').disabled=true;
    try{
      const files=[],used=new Set();progress(3,'Đang chuẩn bị mẫu Word…');
      for(let i=0;i<selected.length;i++){
        const c=selected[i],bytes=await E.generateWord(c.source,c.options,state.workbook.date1904,state.template);
        let base='Bao_cao_tuan_thu_'+E.safeName(c.source.values.C14||c.source.name)+'_'+E.safeName(c.source.values.C12),name=base+'.docx',counter=2;
        while(used.has(name))name=base+'_'+counter+++'.docx';used.add(name);files.push({name,bytes,label:String(c.source.values.C12)});
        progress(5+Math.round((i+1)/selected.length*67),`Đã tạo ${i+1}/${selected.length} báo cáo Word`);await new Promise(r=>setTimeout(r,0));
      }
      progress(76,'Đang thêm sheet TONG_HOP vào cuối Excel…');
      const summary=await E.summaryWorkbook(state.workbook,selected.map(c=>c.source));
      const excelName=E.safeName(state.file.name.replace(/\.xlsx$/i,''))+'_TONG_HOP.xlsx';
      const archive=new JSZip();files.forEach(f=>archive.file('Bao_cao_Word/'+f.name,f.bytes));archive.file(excelName,summary.bytes);
      progress(90,'Đang đóng gói bộ báo cáo…');
      const zip=await archive.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:3}});
      const zipLink=$('download-all');zipLink.href=url(zip,'application/zip');zipLink.download='Bo_bao_cao_tuan_thu_KHCN.zip';
      downloadItem(excelName,'Excel kèm TONG_HOP',summary.bytes,'xlsx');files.forEach(f=>downloadItem(f.name,f.label,f.bytes,'docx'));
      $('result-title').textContent=`${files.length} báo cáo Word và 1 file Excel`;
      progress(100,'Hoàn tất. Chọn file cần tải bên dưới.');reveal('results',true);setSteps(3);$('results').scrollIntoView({behavior:'smooth',block:'start'});
    }catch(error){clearDownloads();$('generate-error').textContent=error.message||'Chưa tạo được báo cáo. Vui lòng kiểm tra lại dữ liệu.';reveal('generate-error',true);reveal('progress-area',false);}
    finally{state.busy=false;$('customer-list').inert=false;$('select-all').disabled=false;$('change-file').disabled=false;$('clear-file').disabled=false;$('choose-word').disabled=false;updateActions();}
  }
  $('dropzone').addEventListener('click',()=>$('file-input').click());$('change-file').addEventListener('click',()=>$('file-input').click());
  $('choose-word').addEventListener('click',()=>$('word-input').click());$('word-input').addEventListener('change',e=>loadWord(e.target.files[0]));
  $('file-input').addEventListener('change',e=>loadFile(e.target.files[0]));$('clear-file').addEventListener('click',reset);
  $('dropzone').addEventListener('dragover',e=>{e.preventDefault();$('dropzone').classList.add('dragging');});
  $('dropzone').addEventListener('dragleave',()=>$('dropzone').classList.remove('dragging'));
  $('dropzone').addEventListener('drop',e=>{e.preventDefault();$('dropzone').classList.remove('dragging');loadFile(e.dataTransfer.files[0]);});
  document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
  $('customer-list').addEventListener('click',e=>{const b=e.target.closest('[data-action="review"]');if(!b||state.busy)return;const c=state.customers[Number(b.dataset.index)];c.open=!c.open;renderCustomers();});
  $('customer-list').addEventListener('change',e=>{
    const el=e.target,c=state.customers[Number(el.dataset.index)];if(!c||state.busy)return;
    if(el.dataset.field==='selected')c.selected=el.checked;else c.options[el.dataset.field]=el.type==='checkbox'?el.checked:el.value;
    clearDownloads();reveal('progress-area',false);reveal('generate-error',false);setSteps(2);renderCustomers();
  });
  $('select-all').addEventListener('click',()=>{if(state.busy)return;const select=!state.customers.every(c=>c.selected);state.customers.forEach(c=>c.selected=select);clearDownloads();reveal('progress-area',false);setSteps(2);renderCustomers();});
  $('generate').addEventListener('click',generate);
  updateTemplate();
  if(globalThis.SAMPLE_WORKBOOK_BASE64){$('sample-download').href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+globalThis.SAMPLE_WORKBOOK_BASE64;reveal('sample-download',true);}
  window.addEventListener('beforeunload',()=>state.urls.forEach(URL.revokeObjectURL));
})();
