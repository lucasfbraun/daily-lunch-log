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