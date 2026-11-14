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
