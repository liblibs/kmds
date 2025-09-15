// JavaScript to handle page loading
function loadPage(page) {
    const contentArea = document.getElementById('page-content');

    fetch(page)
        .then(response => response.text())
        .then(data => {
            contentArea.innerHTML = data;
            // Re-initialize Mermaid after loading new content
            // mermaid.contentLoaded();
        })
        .catch(error => {
            console.error('Error loading the page:', error);
            contentArea.innerHTML = `<p>Sorry, we couldn't load the content. Please try again later.</p>`;
        });
}

// Initial content loading for the homepage
document.addEventListener('DOMContentLoaded', () => {
    loadPage('index.html');
});
