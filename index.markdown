---
# Feel free to add content and custom Front Matter to this file.
# To modify the layout, see https://jekyllrb.com/docs/themes/#overriding-theme-defaults

layout: home
---

<div id="data-container"></div>
<script>
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwjS1UWlP925d5CuJYLG1L02oeNLXFK8EjZR9c1vjj56FIRbqcaGZxcFQcZojikgNdEsA/exec'; // <-- Paste your URL

fetch(WEB_APP_URL)
  .then(response => response.json())
  .then(result => {
    const data = result.data; 
    console.log(data); 
    displayData(data); 
  })
  .catch(error => {
    console.error('Error fetching data:', error);
  });

function displayData(items) {
  const container = document.getElementById('data-container');
  if (!container) return;
  container.innerHTML = '';
  items.forEach(item => {
    const div = document.createElement('div');
    div.innerHTML = `<h3>${item.Name}</h3><p>${item.Description}</p>`;
    container.appendChild(div);
  });
}
</script>
