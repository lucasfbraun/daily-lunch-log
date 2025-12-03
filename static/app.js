document.addEventListener('DOMContentLoaded', ()=>{
  const form = document.getElementById('uploadForm');
  const fileInput = document.getElementById('fileInput');
  const status = document.getElementById('status');
  const downloadArea = document.getElementById('downloadArea');
  const fileDrop = document.querySelector('.file-drop');
  const fileNameEl = document.getElementById('fileName');

  form.addEventListener('submit', async (e)=>{
    e.preventDefault();
    downloadArea.innerHTML='';
    status.textContent='';

    const file = fileInput.files[0];
    if(!file){ status.textContent='Selecione um arquivo .xlsx primeiro.'; return }

    status.textContent='Enviando arquivo...';
    try{
      const formData = new FormData();
      formData.append('file', file);
      const resp = await fetch('/convert',{method:'POST',body:formData});
      if(!resp.ok){
        const txt = await resp.text();
        status.textContent = 'Erro: '+ (txt || resp.statusText);
        return;
      }

      const blob = await resp.blob();
      const filename = resp.headers.get('X-Filename') || 'output.txt';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.id='downloadLink'; a.download = filename; a.textContent = 'Baixar TXT convertido';
      a.className='btn-cta';
      downloadArea.appendChild(a);
      status.textContent = 'Conversão concluída.';
    }catch(err){
      console.error(err);
      status.textContent = 'Erro ao enviar o arquivo.';
    }
  });

  // update filename display when a file is selected
  fileInput.addEventListener('change', ()=>{
    const f = fileInput.files[0];
    if(f){
      fileNameEl.textContent = f.name;
    } else {
      fileNameEl.textContent = '';
    }
  });

  // drag and drop support on the label
  if(fileDrop){
    fileDrop.addEventListener('dragover', (e)=>{ e.preventDefault(); fileDrop.classList.add('dragover'); });
    fileDrop.addEventListener('dragleave', ()=>{ fileDrop.classList.remove('dragover'); });
    fileDrop.addEventListener('drop', (e)=>{
      e.preventDefault(); fileDrop.classList.remove('dragover');
      if(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length){
        fileInput.files = e.dataTransfer.files;
        const evt = new Event('change'); fileInput.dispatchEvent(evt);
      }
    });
  }
});


// --- Novo bloco: Almoços funcionarios PJ ---
document.addEventListener('DOMContentLoaded', ()=>{
  const formPJ = document.getElementById('uploadFormPJ');
  const fileInputPJ = document.getElementById('fileInputPJ');
  const statusPJ = document.getElementById('statusPJ');
  const downloadAreaPJ = document.getElementById('downloadAreaPJ');
  const fileDropPJ = document.querySelector('#fileInputPJ + .file-drop-content')?.parentElement;
  const fileNameElPJ = document.getElementById('fileNamePJ');

  if(!formPJ) return; // Se não existir o formulário PJ, não faz nada

  formPJ.addEventListener('submit', async (e)=>{
    e.preventDefault();
    downloadAreaPJ.innerHTML='';
    statusPJ.textContent='';

    const file = fileInputPJ.files[0];
    if(!file){ statusPJ.textContent='Selecione um arquivo .xlsx primeiro.'; return }

    statusPJ.textContent='Processando arquivo...';
    try{
      const formData = new FormData();
      formData.append('file', file);
      const resp = await fetch('/convert_pj',{method:'POST',body:formData});
      if(!resp.ok){
        const txt = await resp.text();
        statusPJ.textContent = 'Erro: '+ (txt || resp.statusText);
        return;
      }

      const blob = await resp.blob();
      const filename = 'almoco_pj_totalizado.xlsx';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.id='downloadLinkPJ'; a.download = filename; a.textContent = 'Baixar Excel totalizado';
      a.className='btn-cta';
      statusPJ.textContent = 'Totalização concluída.';
      downloadAreaPJ.appendChild(a);
    }catch(err){
      console.error(err);
      statusPJ.textContent = 'Erro ao processar o arquivo.';
    }
  });

  // update filename display when a file is selected
  fileInputPJ.addEventListener('change', ()=>{
    const f = fileInputPJ.files[0];
    if(f){
      fileNameElPJ.textContent = f.name;
    } else {
      fileNameElPJ.textContent = '';
    }
  });

  // drag and drop support
  if(fileDropPJ){
    fileDropPJ.addEventListener('dragover', (e)=>{ e.preventDefault(); fileDropPJ.classList.add('dragover'); });
    fileDropPJ.addEventListener('dragleave', ()=>{ fileDropPJ.classList.remove('dragover'); });
    fileDropPJ.addEventListener('drop', (e)=>{
      e.preventDefault(); fileDropPJ.classList.remove('dragover');
      if(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length){
        fileInputPJ.files = e.dataTransfer.files;
        const evt = new Event('change'); fileInputPJ.dispatchEvent(evt);
      }
    });
  }
});