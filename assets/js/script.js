// JavaScript to handle page loading
function loadPage(page) {
    const contentArea = document.getElementById('page-content');
    const sidebar = document.getElementById('sidebar');

    // Dynamically add the 'active' class to the sidebar
    sidebar.classList.add('active');

    fetch(page)
        .then(response => response.text())
        .then(data => {
            contentArea.innerHTML = data;
        })
        .catch(error => {
            console.error('Error loading the page:', error);
            contentArea.innerHTML = `<p>Sorry, we couldn't load the content. Please try again later.</p>`;
        });
}

// Initial content loading for the homepage
// document.addEventListener('DOMContentLoaded', () => {
//     loadPage('index.html');
// });

function openInNewTab(url) {
  window.open(url, '_blank', 'noopener,noreferrer');
}
