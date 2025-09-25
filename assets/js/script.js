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

document.addEventListener("DOMContentLoaded", () => {
  const copyBtn = document.getElementById("copyBtn");
  const sqlPreview = document.getElementById("sqlPreview");

  // Copy SQL to clipboard
copyBtn.addEventListener("click", () => {
    // Create a temporary textarea
    const textarea = document.createElement("textarea");
    textarea.value = sqlPreview.textContent;
    document.body.appendChild(textarea);

    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);

    // Optional: Give feedback
    copyBtn.textContent = "Copied!";
    setTimeout(() => {
      copyBtn.textContent = "Copy SQL";
    }, 1500);
  });
});
