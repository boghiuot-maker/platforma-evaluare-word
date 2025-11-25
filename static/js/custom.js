document.addEventListener('DOMContentLoaded', function(){
  const send = document.getElementById('send');
  const file = document.getElementById('file');
  const nume = document.getElementById('nume');
  const clasa = document.getElementById('clasa');
  const data_test = document.getElementById('data_test');
  const result = document.getElementById('result');
  send.addEventListener('click', async ()=>{
    if(!nume.value.trim()){ alert('Completează numele!'); return; }
    if(!clasa.value){ alert('Alege clasa!'); return; }
    if(!data_test.value){ alert('Completează data testului!'); return; }
    if(!file.files[0]){ alert('Selectează un fișier .docx'); return; }
    if(!confirm('Ești sigur că acesta este fișierul final pe care vrei să îl trimiți?')) return;
    const fd = new FormData();
    fd.append('file', file.files[0]);
    fd.append('nume', nume.value.trim());
    fd.append('clasa', clasa.value);
    fd.append('data_test', data_test.value);
    result.innerText = 'Se evaluează...';
    try{
      const res = await fetch('/evaluate', {method:'POST', body: fd});
      const data = await res.json();
      if(res.ok){
        let html = `<h3>Scor: ${data.scor} / ${data.punctaj_maxim}</h3>`;
        html += `<p>Număr cuvinte: ${data.nr_cuvinte}</p>`;
        html += '<h4>Feedback</h4><ul>';
        data.feedback.forEach(f=> html += `<li>${f}</li>`);
        html += '</ul>';
        result.innerHTML = html;
        // disable inputs to prevent re-upload
        file.disabled = true;
        send.disabled = true;
      } else {
        result.innerText = data.error || JSON.stringify(data);
      }
    } catch(e){
      result.innerText = 'Eroare: '+e;
    }
  });
});
