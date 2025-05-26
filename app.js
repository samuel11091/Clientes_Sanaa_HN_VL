
document.addEventListener("DOMContentLoaded", () => {
  const root = document.getElementById("root");
  const table = document.createElement("table");
  table.style.borderCollapse = "collapse";
  table.style.width = "100%";
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");
  table.appendChild(thead);
  table.appendChild(tbody);

  const exportButton = document.createElement("button");
  exportButton.textContent = "Exportar a Excel";
  exportButton.style.marginBottom = "1rem";
  exportButton.onclick = () => {
    window.open("clientes_sanaa_template.xlsx", "_blank");
  };

  root.innerHTML = "<h2>Clientes Sanaa</h2>";
  root.appendChild(exportButton);
  root.appendChild(table);

  fetch("clientes_sanaa_template.xlsx")
    .then((res) => res.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (json.length) {
        const headers = json[0];
        const headRow = document.createElement("tr");
        headers.forEach((h) => {
          const th = document.createElement("th");
          th.textContent = h;
          th.style.border = "1px solid #ccc";
          th.style.padding = "8px";
          th.style.backgroundColor = "#f0f0f0";
          headRow.appendChild(th);
        });
        thead.appendChild(headRow);

        for (let i = 1; i < json.length; i++) {
          const row = document.createElement("tr");
          json[i].forEach((cell) => {
            const td = document.createElement("td");
            td.textContent = cell;
            td.style.border = "1px solid #ccc";
            td.style.padding = "6px";
            row.appendChild(td);
          });
          tbody.appendChild(row);
        }
      } else {
        tbody.innerHTML = "<tr><td>No hay datos en el archivo.</td></tr>";
      }
    })
    .catch((err) => {
      root.innerHTML += "<p style='color:red;'>Error al cargar archivo Excel.</p>";
      console.error(err);
    });
});
