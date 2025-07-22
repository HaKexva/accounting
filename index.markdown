---
# Feel free to add content and custom Front Matter to this file.
# To modify the layout, see https://jekyllrb.com/docs/themes/#overriding-theme-defaults

layout: home
---

<div id="data-container"></div>
<script>
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz13j4ymrL3VExn-yYEPsdkEMk9NB7YK5a_2F2lSgWycD_27c--p4h9Uzr5OdFBVZnIxw/exec'; // <-- Paste your URL

fetch(WEB_APP_URL)
  .then(response => response.json())
  .then(result => {
    const data = result.data; 
    console.log(data); 
    // Now use the 'data' array to populate your webpage
    // e.g., build a table, create list items, etc.
    displayData(data); 
  })
  .catch(error => {
    console.error('Error fetching data:', error);
    // Handle errors (e.g., show a message to the user)
  });

function displayData(items) {
  const container = document.getElementById('data-container'); // Assuming you have an element with this ID
  if (!container) return;

  container.innerHTML = ''; // Clear previous content

  items.forEach(item => {
    const div = document.createElement('div');
    // Example: Assuming your sheet has 'Name' and 'Description' columns
    div.innerHTML = `<h3>${item.Name}</h3><p>${item.Description}</p>`; 
    container.appendChild(div);
  });
}
</script>
