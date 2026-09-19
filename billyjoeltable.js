// Path to the Billy Joel media references workbook (same folder)
const xlsxFilePath = 'BillyJoelMediaReferences.xlsx';

// Columns that get a dropdown filter instead of a free-text search box
const SELECT_FILTER_COLUMNS = ['Media'];

// Read the tablesorter theme name straight from the linked theme CSS file
// (e.g. ".../theme.grey.min.css" -> "grey") so swapping that <link> in
// index.html is enough to change the theme - no JS edit required.
function getTablesorterTheme() {
  const link = document.querySelector('link[href*="tablesorter"][href*="theme."]');
  const match = link && link.href.match(/theme\.([a-z0-9_-]+)\.min\.css/i);
  return match ? match[1] : 'default';
}

// Load the XLSX file and initialize the sortable/filterable table
function loadXLSX(url) {
  fetch(encodeURI(url))
    .then(response => {
      if (!response.ok) throw new Error('Network response was not ok');
      return response.arrayBuffer();
    })
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      // raw:false formats dates/durations using the cell's own number format
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
      const headers = jsonData.length ? Object.keys(jsonData[0]) : [];
      return { headers, rows: jsonData };
    })
    .then(({ headers, rows }) => {
      $('#loading').hide();

      // Header row (display label can differ from the underlying xlsx column name)
      const DISPLAY_LABELS = { Title: 'Media Title' };
      const headHtml = headers.map(header => `<th>${DISPLAY_LABELS[header] || header}</th>`).join('');
      $('#table-head-row').html(headHtml);

      // Data rows
      const CELL_CLASSES = {
        Song: 'col-song',
        Title: 'col-title',
        Episode: 'col-episode',
        Timestamp: 'col-timestamp'
      };
      const bodyHtml = rows.map(row => {
        const cells = headers.map(header => {
          const cls = CELL_CLASSES[header] ? ` class="${CELL_CLASSES[header]}"` : '';
          return `<td${cls}>${row[header]}</td>`;
        }).join('');
        return `<tr>${cells}</tr>`;
      }).join('');
      $('#table-body').html(bodyHtml);

      // Build a { columnIndex: true } map so tablesorter renders those
      // columns as a dropdown of unique values instead of a text search box
      const filterFunctions = {};
      headers.forEach((header, index) => {
        if (SELECT_FILTER_COLUMNS.includes(header)) {
          filterFunctions[index] = true;
        }
      });

      $('#myTable').tablesorter({
        theme: getTablesorterTheme(),
        widthFixed: true,
        widgets: ['zebra', 'filter', 'pager'],
        widgetOptions: {
          filter_columnFilters: true,
          filter_functions: filterFunctions,
          filter_placeholder: { search: 'Filter...' },
          filter_searchDelay: 300,
          filter_liveSearch: true,
          filter_reset: '.reset-filters',

          pager_selectors: {
            container: '.pager-controls',
            pageDisplay: '.pagedisplay',
            pageSize: '.pagesize',
            first: '.first',
            previous: '.prev',
            next: '.next',
            last: '.last'
          },
          pager_output: '{startRow} - {endRow} / {filteredRows} rows ({totalRows} total)',
          pager_updateArrows: true,
          pager_startPage: 0,
          pager_pageReset: 0,
          pager_size: 20,
          pager_removeRows: false
        }
      });

      // pager_size above isn't respected on init in this tablesorter build,
      // so force the default page size explicitly right after setup.
      $('#myTable').trigger('pageSize', 20);
    })
    .catch(err => {
      $('#loading').text('Error loading XLSX file: ' + err.message);
    });
}

$(document).ready(() => {
  loadXLSX(xlsxFilePath);
});
