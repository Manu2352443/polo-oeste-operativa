(function () {
    const estados = new WeakMap();

    function valorCelda(texto) {
        const valor = String(texto || "").trim();
        const numero = Number(valor.replace(/\./g, "").replace(",", ".").replace(/[^0-9.-]/g, ""));
        if (valor && !Number.isNaN(numero) && /\d/.test(valor)) return { tipo: "numero", valor: numero };
        return { tipo: "texto", valor: valor.localeCompare ? valor.toLocaleLowerCase("es") : valor.toLowerCase() };
    }

    function actualizarBotones(tabla, estado) {
        tabla.querySelectorAll("thead th[data-sortable]").forEach(function (celda, indice) {
            const boton = celda.querySelector("button.table-sort-button");
            if (!boton) return;
            const activo = estado.columna === indice ? estado.direccion : "";
            boton.dataset.direction = activo;
            boton.setAttribute("aria-label", activo === "asc" ? "Orden ascendente; presiona para invertir" : activo === "desc" ? "Orden descendente; presiona para restaurar" : "Ordenar columna");
        });
    }

    function ordenar(tabla, indice) {
        const cuerpo = tabla.tBodies[0];
        if (!cuerpo) return;
        const estado = estados.get(tabla);
        if (estado.columna !== indice) {
            estado.original = Array.from(cuerpo.rows);
            estado.columna = indice;
            estado.direccion = "asc";
        } else if (estado.direccion === "asc") {
            estado.direccion = "desc";
        } else if (estado.direccion === "desc") {
            estado.columna = null;
            estado.direccion = "";
            estado.original.forEach(function (fila) { cuerpo.appendChild(fila); });
            actualizarBotones(tabla, estado);
            return;
        } else {
            estado.direccion = "asc";
        }

        const filas = Array.from(cuerpo.rows).filter(function (fila) { return fila.cells.length > indice; });
        filas.sort(function (a, b) {
            const valorA = valorCelda(a.cells[indice].textContent);
            const valorB = valorCelda(b.cells[indice].textContent);
            let comparacion = valorA.tipo === "numero" && valorB.tipo === "numero"
                ? valorA.valor - valorB.valor
                : String(valorA.valor).localeCompare(String(valorB.valor), "es", { numeric: true, sensitivity: "base" });
            return estado.direccion === "desc" ? -comparacion : comparacion;
        });
        filas.forEach(function (fila) { cuerpo.appendChild(fila); });
        actualizarBotones(tabla, estado);
    }

    function activar(tabla) {
        if (!tabla || estados.has(tabla)) return;
        const estado = { columna: null, direccion: "", original: Array.from(tabla.tBodies[0] ? tabla.tBodies[0].rows : []) };
        estados.set(tabla, estado);
        tabla.querySelectorAll("thead th").forEach(function (celda, indice) {
            if (!celda.textContent.trim()) return;
            celda.dataset.sortable = "true";
            const etiqueta = celda.textContent.trim();
            celda.textContent = "";
            const boton = document.createElement("button");
            boton.type = "button";
            boton.className = "table-sort-button";
            boton.textContent = etiqueta;
            boton.addEventListener("click", function () { ordenar(tabla, indice); });
            celda.appendChild(boton);
        });
        actualizarBotones(tabla, estado);
    }

    function refrescar(tabla) {
        if (!tabla || !estados.has(tabla)) return;
        const estado = estados.get(tabla);
        estado.columna = null;
        estado.direccion = "";
        estado.original = Array.from(tabla.tBodies[0] ? tabla.tBodies[0].rows : []);
        actualizarBotones(tabla, estado);
    }

    window.PoloOesteTables = {
        activar: activar,
        refrescar: refrescar,
        activarTodo: function () { document.querySelectorAll("table.excel-table").forEach(activar); }
    };
    document.addEventListener("DOMContentLoaded", window.PoloOesteTables.activarTodo);
}());
