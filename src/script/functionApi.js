const rows = [];
var Archivo = "";

document.addEventListener("DOMContentLoaded", () => {
  const botonDescargar = document.getElementById("botonDescargar");
  const tablaContainer = document.getElementById("tablaContainer");
  const btnExcel = document.getElementById("btnExcel");
  const btnenviar = document.getElementById("btnEnviar");
  btnExcel.addEventListener("click", async () => {
    getApiData().then(() => {
      generarArchivoExcel();
    });
  });
  botonDescargar.addEventListener("click", () => {
    getApiData();
  });
  async function ejecutarTodo() {
    const archivoInput = document.getElementById("archivoInput");
    const archivo = archivoInput.files[0];
    try {
      showLoader(); // Mostrar animación de carga
      const [dataDiario, dataComisiones] = await Promise.all([
        getApiData(),
        getApiDataComisiones(),
      ]);
      await Promise.all([
        generarArchivoExcel("miTabla", "Ventas", dataDiario),
        generarArchivoExcel("miTablaComisiones", "Comisiones", dataComisiones),
      ]);
      enviarCorreo(
        "edisonnacional1@hotmail.com",
        "Facuta electronica",
        "Contenido del correo",
        archivo
      );
      hideLoader(); // Ocultar animación de carga
    } catch (error) {
      console.error("Ocurrió un error:", error);
      hideLoader(); // Ocultar animación de carga en caso de error
    }
  }

  btnenviar.addEventListener("click", () => {
    ejecutarTodo();
  });
  ejecutarTodo();
  function getApiData() {
    // Hacer la solicitud HTTP a la API
    return fetch("http://localhost:3000/api/v1/resumen/diario")
      .then((response) => {
        if (!response.ok) {
          throw new Error("La solicitud no se pudo completar");
        }
        return response.json();
      })
      .then((data) => {
        // Crear y mostrar la tabla

        const tabla = document.createElement("table");
        tabla.id = "miTabla";
        const encabezados = Object.keys(data.datos[0]);

        // Crear fila de encabezados
        const encabezadosRow = tabla.insertRow();
        encabezados.forEach((encabezado) => {
          rows.push(encabezado);

          const th = document.createElement("th");
          th.textContent = encabezado;

          encabezadosRow.appendChild(th);
        });

        // Crear filas de datos
        data.datos.forEach((item) => {
          const fila = tabla.insertRow();
          encabezados.forEach((encabezado) => {
            const td = fila.insertCell();
            if (encabezado !== "Almacen") {
              // Formatear el valor a un formato decimal con 18 dígitos decimales y 0 dígitos después del punto decimal
              const valorFormateado = parseFloat(item[encabezado]).toFixed(2);
              td.textContent = valorFormateado;
            } else {
              td.textContent = item[encabezado];
            }
          });
        });

        const sumaRow = tabla.insertRow();
        encabezados.forEach((encabezado) => {
          const td = sumaRow.insertCell();
          if (encabezado !== "Almacen" && !isNaN(data.datos[0][encabezado])) {
            const sumaColumna = data.datos.reduce(
              (total, item) => total + parseFloat(item[encabezado]),
              0
            );
            td.textContent = sumaColumna.toFixed(2);
          } else {
            td.textContent = "";
          }
        });

        // Mostrar la tabla
        tablaContainer.innerHTML = "";
        tablaContainer.appendChild(tabla);
      })
      .catch((error) => {
        console.error("Error al obtener los datos:", error);
        alert("Error al obtener los datos");
      });
  }

  function getApiDataComisiones() {
    // Hacer la solicitud HTTP a la API
    return fetch("http://localhost:3000/api/v1/resumen/comisiones")
      .then((response) => {
        if (!response.ok) {
          throw new Error("La solicitud no se pudo completar");
        }
        return response.json();
      })
      .then((data) => {
        // Crear la tabla
        let tablaHTML = '<table id="miTablaComisiones">';
        const encabezados = Object.keys(data.datos[0]);

        // Crear fila de encabezados
        tablaHTML += "<tr>";
        encabezados.forEach((encabezado) => {
          tablaHTML += "<th>" + encabezado + "</th>";
        });
        tablaHTML += "</tr>";

        // Crear filas de datos
        data.datos.forEach((item) => {
          tablaHTML += "<tr>";
          encabezados.forEach((encabezado) => {
            tablaHTML += "<td>";
            if (
              encabezado !== "Almacen" &&
              encabezado !== "Vendedor" &&
              encabezado !== "Cargo"
            ) {
              const valorFormateado = parseFloat(item[encabezado]).toFixed(2);
              tablaHTML += valorFormateado;
            } else {
              tablaHTML += item[encabezado];
            }
            tablaHTML += "</td>";
          });
          tablaHTML += "</tr>";
        });

        // Calcular y agregar fila de suma
        tablaHTML += "<tr>";
        encabezados.forEach((encabezado) => {
          tablaHTML += "<td>";
          if (encabezado !== "Almacen" && !isNaN(data.datos[0][encabezado])) {
            const sumaColumna = data.datos.reduce(
              (total, item) => total + parseFloat(item[encabezado]),
              0
            );
            tablaHTML += sumaColumna.toFixed(2);
          } else {
            tablaHTML += "";
          }
          tablaHTML += "</td>";
        });
        tablaHTML += "</tr>";

        // Cerrar la tabla
        tablaHTML += "</table>";

        // Mostrar la tabla dentro del div existente
        const tablaContainer = document.getElementById("tablaComisiones");
        if (tablaContainer) {
          tablaContainer.innerHTML = tablaHTML;
        } else {
          console.error("No se encontró el contenedor de la tabla.");
        }
      })
      .catch((error) => {
        console.error("Error al obtener los datos comisiones:", error);
        alert("Error al obtener los datos");
      });
  }

  function generarArchivoExcel(nameTable, nameFile) {
    const table = document.getElementById(nameTable);

    // Crear un libro de trabajo Excel
    const workbook = XLSX.utils.book_new();

    // Obtener los datos de la tabla
    const sheetData = [];
    for (let i = 0; i < table.rows.length; i++) {
      const rowData = [];
      const row = table.rows[i];
      for (let j = 0; j < row.cells.length; j++) {
        rowData.push(row.cells[j].textContent);
      }
      sheetData.push(rowData);
    }

    // Crear una hoja de trabajo y convertir los datos de la tabla a formato de hoja de cálculo
    const sheet = XLSX.utils.aoa_to_sheet(sheetData);

    // Crear un estilo personalizado para las celdas
    const style = {
      font: { name: "Arial", sz: 12 },
      alignment: { horizontal: "left", vertical: "middle" },
      fill: { fgColor: { rgb: "FFFF00" } }, // Color de fondo amarillo
    };

    // Aplicar el estilo a las celdas
    for (let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
      for (let colIndex = 0; colIndex < sheetData[0].length; colIndex++) {
        const cellAddress = XLSX.utils.encode_cell({
          r: rowIndex,
          c: colIndex,
        });
        if (!sheet[cellAddress]) {
          sheet[cellAddress] = {};
        }
        sheet[cellAddress].s = style;
      }
    }

    // Agregar la hoja de trabajo al libro de trabajo
    XLSX.utils.book_append_sheet(workbook, sheet, "Hoja1");

    // Crear un archivo Excel binario a partir del libro de trabajo
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    // Convertir el archivo Excel binario a un Blob
    const blob = new Blob([excelBuffer], { type: "application/octet-stream" });

    const currentDate = new Date();
    currentDate.setDate(currentDate.getDate());
    const formattedDate = currentDate
      .toISOString()
      .slice(0, 10)
      .replace(/-/g, "");
    const year = currentDate.getFullYear();
    const month = ("0" + (currentDate.getMonth() + 1)).slice(-2);
    const day = ("0" + currentDate.getDate()).slice(-2);

    // Construir el nombre del archivo con la fecha actual
    const fileName = `${nameFile}_${year}_${month}_${day}.xlsx`;
    Archivo = fileName;
    // Descargar el archivo Excel
    saveAs(blob, fileName);
  }

  function enviarCorreo(destinatario, asunto, contenido, archivo) {
    if (!destinatario || !asunto || !contenido) {
      console.error("Todos los campos son obligatorios");
      return;
    }
    const formData = new FormData();
    formData.append("destinatario", destinatario);
    formData.append("asunto", asunto);
    formData.append("contenido", contenido);
    formData.append("adjunto", archivo);
    console.log(formData.get("destinatario"));

    fetch("http://localhost:3000/api/v1/resumen/comisiones/email", {
      method: "POST",
      body: formData,
    })
      .then((response) => {
        if (response.ok) {
          console.log("Correo electrónico enviado correctamente");
        } else {
          console.error("Error al enviar el correo electrónico");
        }
      })
      .catch((error) => {
        console.error("Error al enviar el correo electrónico:", error);
      });
  }
  function showLoader() {
    loader.style.display = "block";
  }

  function hideLoader() {
    loader.style.display = "none";
  }

  // Ejemplo de uso
  // enviarCorreo('destinatario@example.com', 'Asunto del correo', 'Contenido del correo');
});
