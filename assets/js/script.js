// JavaScript to handle page loading
// function loadPage(page) {
//     const contentArea = document.getElementById('page-content');
//     const sidebar = document.getElementById('sidebar');

//     // Dynamically add the 'active' class to the sidebar
//     sidebar.classList.add('active');

//     fetch(page)
//         .then(response => response.text())
//         .then(data => {
//             contentArea.innerHTML = data;
//         })
//         .catch(error => {
//             console.error('Error loading the page:', error);
//             contentArea.innerHTML = `<p>Sorry, we couldn't load the content. Please try again later.</p>`;
//         }); 
// }

// Keep this at file scope if you want to cancel a previous load when user clicks fast
let _currentFetchController = null;

function loadPage(page) {
  const contentArea = document.getElementById('page-content');
  const sidebar = document.getElementById('sidebar');

  // Mark sidebar as active (as you had)
  sidebar.classList.add('active');

  // Show a quick loading state
  contentArea.innerHTML = '<p>Loading…</p>';

  // Abort any in-flight request (user clicked multiple links quickly)
  if (_currentFetchController) {
    try { _currentFetchController.abort(); } catch (_) {}
  }
  _currentFetchController = new AbortController();
  const { signal } = _currentFetchController;

  // Optional: disable cache if your templates change often
  fetch(page, { cache: 'no-cache', signal })
    .then(res => {
      if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
      return res.text();
    })
    .then(html => {
      // Inject the HTML
      contentArea.innerHTML = html;

      // IMPORTANT: highlight after injecting
      if (window.SqlFormatterLite && typeof window.SqlFormatterLite.refresh === 'function') {
        window.SqlFormatterLite.refresh();
      } else {
        // This means sql-highlighter.js didn't load yet or loaded after this file.
        console.warn(
          'SqlFormatterLite is not available. ' +
          'Ensure sql-highlighter.js is loaded before this script and with defer.'
        );
      }
    })
    .catch(err => {
      // Ignore AbortError (user navigated again quickly)
      if (err && err.name === 'AbortError') return;
      console.error('Error loading the page:', err);
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


