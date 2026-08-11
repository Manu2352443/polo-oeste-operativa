/* Abre el calendario nativo del navegador desde todo el campo de fecha.
   La interfaz nativa conserva accesibilidad, teclado y compatibilidad móvil. */
document.querySelectorAll('input[type="date"], input[type="month"]').forEach(function (campo) {
    campo.classList.add("date-picker-input");
    campo.setAttribute("title", "Abrir calendario");

    campo.addEventListener("click", function () {
        if (typeof campo.showPicker === "function") {
            try { campo.showPicker(); } catch (error) { /* El navegador lo abre por defecto. */ }
        }
    });
});
