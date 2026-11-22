// site.js - three small helpers. No outside libraries.
(function () {
  // 1) Show each game's start time in the visitor's own time zone.
  document.querySelectorAll('time[data-local]').forEach(function (el) {
    var d = new Date(el.getAttribute('datetime'));
    if (isNaN(d)) return;
    el.textContent = d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit', timeZoneName: 'short' });
  });

  // 2) Click a column heading to sort the table by that column.
  document.querySelectorAll('table.sortable').forEach(function (table) {
    var tbody = table.tBodies[0];
    table.querySelectorAll('th[data-sort]').forEach(function (th) {
      var button = th.querySelector('button');
      if (!button) return;
      button.addEventListener('click', function () {
        var column = th.cellIndex;
        var numeric = th.getAttribute('data-sort') === 'num';
        var firstDirection = th.getAttribute('data-first') || (numeric ? 'descending' : 'ascending');
        var current = th.getAttribute('aria-sort');
        var direction = current ? (current === 'ascending' ? 'descending' : 'ascending') : firstDirection;

        table.querySelectorAll('th').forEach(function (other) { other.removeAttribute('aria-sort'); });
        th.setAttribute('aria-sort', direction);

        var rows = Array.prototype.slice.call(tbody.rows);
        rows.sort(function (a, b) {
          var x = a.cells[column].getAttribute('data-v');
          var y = b.cells[column].getAttribute('data-v');
          var result;
          if (numeric) {
            x = parseFloat(x); y = parseFloat(y);
            if (isNaN(x)) x = -Infinity;
            if (isNaN(y)) y = -Infinity;
            result = x - y;
          } else {
            result = String(x).localeCompare(String(y));
          }
          return direction === 'ascending' ? result : -result;
        });
        rows.forEach(function (row) { tbody.appendChild(row); });
      });
    });
  });

  // 4) Rule rows that open a panel under the table ("which teams fit this rule right now"). One panel open at a time.
  // Without this script every list simply stays visible under the table, so nothing is ever hidden behind a dead button.
  // The 'js' class is what hides the lists until a row is clicked.
  // Panels are opened and closed with BOTH a class and the 'hidden' attribute, so this works whichever version of the
  // page or stylesheet a browser has cached.
  var ruleButtons = document.querySelectorAll('button.rule-toggle');
  if (ruleButtons.length) document.documentElement.classList.add('js');
  function setPanel(panel, open) {
    if (!panel) return;
    panel.hidden = !open;
    panel.classList.toggle('open', open);
  }
  ruleButtons.forEach(function (button) { setPanel(document.getElementById(button.getAttribute('data-panel')), false); });
  ruleButtons.forEach(function (button) {
    var row = button.closest('tr');
    function toggle() {
      var panel = document.getElementById(button.getAttribute('data-panel'));
      var wasOpen = button.getAttribute('aria-expanded') === 'true';
      ruleButtons.forEach(function (other) {                       // close everything first
        other.setAttribute('aria-expanded', 'false');
        other.closest('tr').classList.remove('open');
        setPanel(document.getElementById(other.getAttribute('data-panel')), false);
      });
      if (!wasOpen && panel) {                                    // then open this one (clicking an open row closes it)
        button.setAttribute('aria-expanded', 'true');
        row.classList.add('open');
        setPanel(panel, true);
        if (panel.scrollIntoView) panel.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }
    }
    button.addEventListener('click', function (event) { event.stopPropagation(); toggle(); });
    row.addEventListener('click', toggle);                        // clicking anywhere on the row works too
  });

  // 5) Game cards on the home page that open a matchup panel under tonight's games. One at a time; clicking anywhere on
  // the card works, and clicking an open card closes it. Without this script every panel simply stays visible.
  var gameButtons = document.querySelectorAll('button.game-toggle');
  if (gameButtons.length) document.documentElement.classList.add('js');
  gameButtons.forEach(function (button) { setPanel(document.getElementById(button.getAttribute('data-panel')), false); });
  gameButtons.forEach(function (button) {
    var card = button.closest('.game');
    function toggle() {
      var panel = document.getElementById(button.getAttribute('data-panel'));
      var wasOpen = button.getAttribute('aria-expanded') === 'true';
      gameButtons.forEach(function (other) {
        other.setAttribute('aria-expanded', 'false');
        other.closest('.game').classList.remove('open');
        setPanel(document.getElementById(other.getAttribute('data-panel')), false);
      });
      if (!wasOpen && panel) {
        button.setAttribute('aria-expanded', 'true');
        card.classList.add('open');
        setPanel(panel, true);
        if (panel.scrollIntoView) panel.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }
    }
    button.addEventListener('click', function (event) { event.stopPropagation(); toggle(); });
    card.addEventListener('click', toggle);
  });

  // 3) Filter buttons (All / Eastern / Western) above a table.
  document.querySelectorAll('[data-filter-group]').forEach(function (group) {
    var table = document.getElementById(group.getAttribute('data-filter-group'));
    if (!table) return;
    group.querySelectorAll('button').forEach(function (button) {
      button.addEventListener('click', function () {
        var wanted = button.getAttribute('data-filter');
        group.querySelectorAll('button').forEach(function (other) {
          other.setAttribute('aria-pressed', other === button ? 'true' : 'false');
        });
        table.querySelectorAll('tbody tr').forEach(function (row) {
          row.hidden = !(wanted === 'all' || row.getAttribute('data-conf') === wanted);
        });
      });
    });
  });
})();
