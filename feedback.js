(function () {
    function obtenerTema() {
        try { return localStorage.getItem("polo-oeste-theme") || "dark"; }
        catch (error) { return "dark"; }
    }

    function aplicarTema(tema) {
        const valor = tema === "light" ? "light" : "dark";
        document.documentElement.dataset.theme = valor;
        try { localStorage.setItem("polo-oeste-theme", valor); } catch (error) { /* almacenamiento no disponible */ }
        document.dispatchEvent(new CustomEvent("polo-oeste-theme", { detail: { tema: valor } }));
        return valor;
    }

    aplicarTema(obtenerTema());
    function isError(message) {
        return /error|incorrect|inválid|inval|fall|no se pudo|debe|obligator/i.test(String(message || ""));
    }

    function showFeedback(message, type) {
        if (!message || !document.body) return;
        const previous = document.querySelector(".po-feedback");
        if (previous) previous.remove();
        const error = type === "error" || (type !== "success" && isError(message));
        const layer = document.createElement("div");
        layer.className = "po-feedback" + (error ? " po-feedback--error" : "");
        layer.innerHTML = '<div class="po-feedback__box" role="status" aria-live="polite"><div class="po-feedback__icon">' + (error ? "×" : "✓") + '</div><div class="po-feedback__text"></div></div>';
        layer.querySelector(".po-feedback__text").textContent = message;
        document.body.appendChild(layer);
        window.setTimeout(() => layer.remove(), 2050);
    }

    function showProgress(text, detail) {
        const previous = document.querySelector(".po-progress");
        if (previous) previous.remove();
        const layer = document.createElement("div");
        layer.className = "po-progress";
        layer.innerHTML = '<div class="po-progress__box" role="status" aria-live="polite"><div class="po-progress__spinner"></div><div><div class="po-progress__text"></div><span class="po-progress__detail"></span></div></div>';
        layer.querySelector(".po-progress__text").textContent = text || "Procesando…";
        layer.querySelector(".po-progress__detail").textContent = detail || "Espera un momento";
        document.body.appendChild(layer);
    }

    window.alert = function (message) { showFeedback(message); };
    window.PoloOeste = { showFeedback, showProgress, setTheme: aplicarTema, getTheme: obtenerTema };

    document.addEventListener("DOMContentLoaded", function () {
        const params = new URLSearchParams(window.location.search);
        let message = params.get("mensaje");
        let type = "success";
        if (params.get("login") === "ok") message = "Sesión iniciada";
        if (params.get("login") === "error") { message = "Credenciales erróneas"; type = "error"; }
        if (params.get("logout") === "ok") message = "Sesión cerrada";
        if (message) {
            showFeedback(message, type);
            ["mensaje", "login", "logout"].forEach((key) => params.delete(key));
            history.replaceState({}, "", location.pathname + (params.toString() ? "?" + params : "") + location.hash);
        }
        document.querySelectorAll("form[data-progress]").forEach(function (form) {
            form.addEventListener("submit", function () {
                if (form.checkValidity()) showProgress(form.dataset.progress, form.dataset.progressDetail);
            });
        });
    });
}());
