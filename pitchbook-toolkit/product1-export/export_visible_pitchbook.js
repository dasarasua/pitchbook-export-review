(function() {
  // Seleccionamos los ROWS
  const rows = document.querySelectorAll('.data-table__row');

  // Obtenemos las cabeceras de la parte scrollable
  const headerCells = document.querySelectorAll('.data-table__header-cell');
  const header = ['"Company Name"']; // Añadimos primero columna de Companies

  headerCells.forEach(cell => {
    const text = cell.innerText.trim();
    header.push('"' + text.replace(/"/g, '""') + '"');
  });

  const csv = [];
  csv.push(header.join(','));

  // Recorremos las filas
  rows.forEach(row => {
    const rowData = [];

    // 1️⃣ Obtenemos el nombre de la Company (columna fija)
    const companyName = row.querySelector('.data-table__cell-0 .ellipsis');
    const companyText = companyName ? companyName.innerText.trim() : '';
    rowData.push('"' + companyText.replace(/"/g, '""') + '"');

    // 2️⃣ Obtenemos el resto de columnas (scrollable)
    const cells = row.querySelectorAll('.data-table__cell:not(.data-table__cell-0)');

    cells.forEach(cell => {
      let textSpan = cell.querySelector('.ellipsis');
      let text = '';

      if (textSpan) {
        text = textSpan.innerText.trim();
      } else {
        text = cell.innerText.trim();
      }

      rowData.push('"' + text.replace(/"/g, '""') + '"');
    });

    if (rowData.length > 0 && rowData.some(val => val !== '""')) {
      csv.push(rowData.join(','));
    }
  });

  // Generamos el CSV
  const csvString = csv.join('\n');
  const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });

  // Descargar como archivo
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', 'pitchbook_export_clean.csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
})();
