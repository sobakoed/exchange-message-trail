/*
  Exchange Message Trail: the script of the user guide.
  1. The button "Theme" changes the theme: system, dark, light. The browser keeps the selection.
  2. The table of contents shows the section that is on the screen.
  3. A click on a screenshot shows it in its full size.
  4. The button "Copy" copies a text, for example the wallet address.
  The browser can block the storage (a local file). The page operates also without it.
*/
(function () {
  var KEY = 'emt-help-theme';
  var order = ['system', 'dark', 'light'];

  function readTheme() { try { return localStorage.getItem(KEY) || 'system'; } catch (e) { return 'system'; } }
  function writeTheme(v) { try { localStorage.setItem(KEY, v); } catch (e) { } }
  function apply(v) {
    if (v === 'system') { document.documentElement.removeAttribute('data-theme'); }
    else { document.documentElement.setAttribute('data-theme', v); }
    var b = document.getElementById('theme-button');
    if (b) { b.textContent = b.getAttribute('data-' + v) || v; }
  }

  apply(readTheme());
  document.addEventListener('DOMContentLoaded', function () {
    apply(readTheme());
    var b = document.getElementById('theme-button');
    if (b) {
      b.addEventListener('click', function () {
        var next = order[(order.indexOf(readTheme()) + 1) % order.length];
        writeTheme(next); apply(next);
      });
    }

    // The table of contents: the link of the section on the screen is active.
    var links = Array.prototype.slice.call(document.querySelectorAll('nav.toc a[href^="#"]'));
    var targets = links.map(function (a) { return document.getElementById(a.getAttribute('href').slice(1)); });
    function mark() {
      var y = window.scrollY + 90, current = 0;
      for (var i = 0; i < targets.length; i++) { if (targets[i] && targets[i].offsetTop <= y) { current = i; } }
      links.forEach(function (a, i) { a.classList.toggle('active', i === current); });
    }
    window.addEventListener('scroll', mark, { passive: true });
    mark();

    // 3. A click on a screenshot shows it in its full size. A click, the key Esc or the button "x" closes it.
    var box = document.createElement('div');
    box.className = 'zoom';
    box.innerHTML = '<button type="button" class="zoom-close" aria-label="Close">&times;</button><img alt="">';
    document.body.appendChild(box);
    var big = box.querySelector('img');
    function closeZoom() { box.classList.remove('open'); document.body.classList.remove('zoom-on'); big.removeAttribute('src'); }
    function openZoom(img) {
      big.src = img.currentSrc || img.src;
      big.alt = img.alt;
      box.classList.add('open');
      document.body.classList.add('zoom-on');
      box.scrollTop = 0; box.scrollLeft = 0;
    }
    Array.prototype.forEach.call(document.querySelectorAll('figure img'), function (img) {
      img.classList.add('zoomable');
      img.title = img.alt;
      img.addEventListener('click', function () { openZoom(img); });
    });
    box.addEventListener('click', closeZoom);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && box.classList.contains('open')) { closeZoom(); } });

    // 4. The button "Copy" copies the text of an element. Without the clipboard API, the text is selected (Ctrl+C copies it).
    Array.prototype.forEach.call(document.querySelectorAll('button.copy'), function (b) {
      b.addEventListener('click', function () {
        var el = document.getElementById(b.getAttribute('data-copy'));
        if (!el) { return; }
        var text = el.textContent.trim(), label = b.textContent;
        function done() { b.textContent = b.getAttribute('data-done') || label; b.classList.add('done'); setTimeout(function () { b.textContent = label; b.classList.remove('done'); }, 2000); }
        function select() { var r = document.createRange(); r.selectNodeContents(el); var s = window.getSelection(); s.removeAllRanges(); s.addRange(r); }
        if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(text).then(done, select); }
        else { select(); try { if (document.execCommand('copy')) { done(); } } catch (e) { } }
      });
    });  });
})();
