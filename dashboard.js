const operariosActivos = document.getElementById("operariosActivos");
const sesionesPausadas = document.getElementById("sesionesPausadas");
const ultimaActualizacion = document.getElementById("ultimaActualizacion");
const tituloCalendario = document.getElementById("tituloCalendario");
const calendario = document.getElementById("calendario");
const embarquesProceso = document.getElementById("embarquesProceso");
const embarquesProximos = document.getElementById("embarquesProximos");
const mesAnterior = document.getElementById("mesAnterior");
const mesSiguiente = document.getElementById("mesSiguiente");
const tituloResumenAnterior = document.getElementById("tituloResumenAnterior");
const resumenPicking = document.getElementById("resumenPicking");
const resumenAlmacenaje = document.getElementById("resumenAlmacenaje");
const resumenExpedicion = document.getElementById("resumenExpedicion");
const fechaPicking = document.getElementById("fechaPicking");
const fechaAlmacenaje = document.getElementById("fechaAlmacenaje");
const fechaExpedicion = document.getElementById("fechaExpedicion");
const resumenFuncionarios = document.getElementById("resumenFuncionarios");
const enlaceResumenMetricas = document.getElementById("enlaceResumenMetricas");
const alertaEmbarques = document.getElementById("alertaEmbarques");
const tituloDistribucion = document.getElementById("tituloDistribucion");
const graficaDistribucion = document.getElementById("graficaDistribucion");
const totalDistribucion = document.getElementById("totalDistribucion");
const leyendaDistribucion = document.getElementById("leyendaDistribucion");
const funcionariosDisponibles = document.getElementById("funcionariosDisponibles");
const detalleDisponibilidad = document.getElementById("detalleDisponibilidad");
const tareasHoy = document.getElementById("tareasHoy");

const fechaActual = new Date();

let anioVisible = fechaActual.getFullYear();
let mesVisible = fechaActual.getMonth();
let ingresosCalendario = [];


function fechaLocal(fecha) {
    const anio = fecha.getFullYear();
    const mes = String(fecha.getMonth() + 1).padStart(2, "0");
    const dia = String(fecha.getDate()).padStart(2, "0");

    return `${anio}-${mes}-${dia}`;
}


function mostrarCalendario() {
    const nombreMes = new Date(anioVisible, mesVisible, 1)
        .toLocaleDateString("es-UY", {
            month: "long",
            year: "numeric"
        });

    tituloCalendario.textContent =
        nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    const primerDia = new Date(anioVisible, mesVisible, 1);
    const ultimoDia = new Date(anioVisible, mesVisible + 1, 0);

    let inicioSemana = primerDia.getDay() - 1;

    if (inicioSemana < 0) {
        inicioSemana = 6;
    }

    calendario.innerHTML = "";

    for (let i = 0; i < inicioSemana; i += 1) {
        const vacio = document.createElement("span");
        vacio.className = "calendar-day empty";
        calendario.appendChild(vacio);
    }

    for (let dia = 1; dia <= ultimoDia.getDate(); dia += 1) {
        const celda = document.createElement("span");

        const fechaDia =
            `${anioVisible}-${String(mesVisible + 1).padStart(2, "0")}-` +
            String(dia).padStart(2, "0");

        const ingresosDelDia = ingresosCalendario.filter(function (ingreso) {
            return ingreso.fecha === fechaDia;
        });

        celda.className = "calendar-day";
        celda.textContent = dia;

        if (
            dia === fechaActual.getDate() &&
            mesVisible === fechaActual.getMonth() &&
            anioVisible === fechaActual.getFullYear()
        ) {
            celda.classList.add("today");
        }

        if (ingresosDelDia.length > 0) {
            const confirmado = ingresosDelDia.some(function (ingreso) {
                return ingreso.estado !== "Proximo";
            });

            const vencido = ingresosDelDia.some(function (ingreso) {
                return (
                    ingreso.estado === "Proximo" &&
                    ingreso.fecha < fechaLocal(fechaActual)
                );
            });

            if (confirmado) {
                celda.classList.add("ingreso-confirmado");
            } else if (vencido) {
                celda.classList.add("ingreso-vencido");
            } else {
                celda.classList.add("ingreso-pendiente");
            }

            celda.title = ingresosDelDia
                .map(function (ingreso) {
                    return `${ingreso.proveedor} - ${ingreso.estado}`;
                })
                .join("\n");
        }

        calendario.appendChild(celda);
    }
}


function actualizarHora() {
    const ahora = new Date();

    ultimaActualizacion.textContent =
        "Actualizado " +
        ahora.toLocaleTimeString("es-UY", {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit"
        });
}


async function actualizarResumenActividad() {
    try {
        const respuesta = await fetch("/api/actividad/registros", {
            cache: "no-store"
        });

        const datos = await respuesta.json();

        const activos = datos.registros.filter(function (registro) {
            return registro.status === "En operacion (Activo)";
        });

        const pausados = datos.registros.filter(function (registro) {
            return registro.status === "En operacion (Pausado)";
        });

        operariosActivos.textContent = activos.length;
        sesionesPausadas.textContent = pausados.length;
        actualizarHora();
    } catch (error) {
        operariosActivos.textContent = "-";
        sesionesPausadas.textContent = "-";
    }
}


async function actualizarResumenEmbarques() {
    try {
        const respuesta = await fetch("/api/embarques/resumen", {
            cache: "no-store"
        });

        const datos = await respuesta.json();

        embarquesProceso.textContent = datos.en_proceso;
        embarquesProximos.textContent = datos.proximos;

        ingresosCalendario = datos.ingresos;
        mostrarCalendario();
        actualizarAlertaEmbarques(datos.ingresos);
    } catch (error) {
        embarquesProceso.textContent = "-";
        embarquesProximos.textContent = "-";
    }
}


async function actualizarDisponibilidadPersonal() {
    try {
        const respuesta = await fetch("/api/panel/disponibilidad", { cache: "no-store" });
        if (!respuesta.ok) throw new Error();
        const datos = await respuesta.json();
        funcionariosDisponibles.textContent = datos.disponibles;
        detalleDisponibilidad.textContent = datos.total
            ? datos.disponibles + " de " + datos.total + " en turno"
            : "Sin funcionarios";
        detalleDisponibilidad.title = (datos.nombres || []).join(", ") || "No hay funcionarios dentro de horario.";
    } catch (error) {
        funcionariosDisponibles.textContent = "-";
        detalleDisponibilidad.textContent = "No disponible";
    }
}


async function actualizarTareasHoy() {
    try {
        const respuesta = await fetch("/api/panel/tareas-hoy", { cache: "no-store" });
        if (!respuesta.ok) throw new Error();
        const datos = await respuesta.json();
        tareasHoy.innerHTML = "";
        if (!datos.tareas.length) {
            tareasHoy.innerHTML = "<p>No tienes pendientes para hoy.</p>";
            return;
        }
        datos.tareas.forEach(function (tarea) {
            const fila = document.createElement("div");
            const prioridad = tarea.prioridad === "Alta" ? "high" : tarea.prioridad === "Baja" ? "low" : "";
            fila.className = "today-task " + prioridad;
            const titulo = document.createElement("span");
            titulo.textContent = tarea.titulo;
            const detalle = document.createElement("small");
            detalle.textContent = tarea.hora || tarea.prioridad;
            fila.appendChild(titulo);
            fila.appendChild(detalle);
            tareasHoy.appendChild(fila);
        });
    } catch (error) {
        tareasHoy.innerHTML = "<p>No se pudo consultar la agenda.</p>";
    }
}


function actualizarAlertaEmbarques(ingresos) {
    const hoy = fechaLocal(fechaActual);
    const manana = new Date(fechaActual);
    manana.setDate(manana.getDate() + 1);
    const fechaManana = fechaLocal(manana);
    const cercanos = (ingresos || []).filter(function (ingreso) {
        return ingreso.estado === "Proximo" && (ingreso.fecha === hoy || ingreso.fecha === fechaManana);
    });
    if (!cercanos.length) {
        alertaEmbarques.hidden = true;
        return;
    }
    const hoyCantidad = cercanos.filter(function (ingreso) { return ingreso.fecha === hoy; }).length;
    const mananaCantidad = cercanos.length - hoyCantidad;
    const partes = [];
    if (hoyCantidad) partes.push(hoyCantidad + (hoyCantidad === 1 ? " ingreso hoy" : " ingresos hoy"));
    if (mananaCantidad) partes.push(mananaCantidad + (mananaCantidad === 1 ? " ingreso mañana" : " ingresos mañana"));
    alertaEmbarques.textContent = partes.join(" · ");
    alertaEmbarques.hidden = false;
}


async function actualizarResumenOperativoAnterior() {
    try {
        const respuesta = await fetch("/api/metricas/resumen-anterior", {
            cache: "no-store"
        });
        if (!respuesta.ok) throw new Error();

        const datos = await respuesta.json();
        function etiquetaFecha(fecha) {
            if (!fecha) return "Sin registros";
            return "Último: " + new Date(fecha + "T12:00:00").toLocaleDateString("es-UY", {
                day: "2-digit", month: "2-digit", year: "numeric"
            });
        }

        tituloResumenAnterior.textContent = datos.fecha_general
            ? "Última actualización: " + new Date(datos.fecha_general + "T12:00:00").toLocaleDateString("es-UY", { day: "2-digit", month: "long" })
            : "Últimos datos operativos";
        resumenPicking.textContent = datos.picking.unidades;
        resumenAlmacenaje.textContent = datos.almacenaje.unidades;
        resumenExpedicion.textContent = datos.expedicion.unidades;
        fechaPicking.textContent = etiquetaFecha(datos.picking.fecha);
        fechaAlmacenaje.textContent = etiquetaFecha(datos.almacenaje.fecha);
        fechaExpedicion.textContent = etiquetaFecha(datos.expedicion.fecha);
        resumenFuncionarios.textContent = "Cada valor corresponde al último día con datos de su tarea.";
        enlaceResumenMetricas.href = datos.fecha_general
            ? "/metricas?mes=" + encodeURIComponent(datos.fecha_general.slice(0, 7))
            : "/metricas";
        actualizarDistribucion(datos);
    } catch (error) {
        tituloResumenAnterior.textContent = "Últimos datos operativos";
        resumenFuncionarios.textContent = "No se pudo consultar el resumen operativo.";
    }
}


function actualizarDistribucion(datos) {
    const configuracion = [
        { nombre: "Picking", color: "#8fd269" },
        { nombre: "Almacenaje", color: "#aebcad" },
        { nombre: "Expedición", color: "#e0c56d" }
    ];
    const distribucion = datos.distribucion || {};
    const total = configuracion.reduce(function (acumulado, tarea) {
        return acumulado + Number(distribucion[tarea.nombre] || 0);
    }, 0);

    totalDistribucion.textContent = total.toLocaleString("es-UY");
    tituloDistribucion.textContent = datos.fecha_distribucion
        ? "Distribución: " + new Date(datos.fecha_distribucion + "T12:00:00").toLocaleDateString("es-UY", { day: "2-digit", month: "long", year: "numeric" })
        : "Último día operativo";

    if (!total) {
        graficaDistribucion.style.background = "conic-gradient(#304035 0 100%)";
        leyendaDistribucion.innerHTML = '<p class="empty-share">Aún no hay unidades para representar.</p>';
        return;
    }

    let acumulado = 0;
    const segmentos = configuracion.map(function (tarea) {
        const unidades = Number(distribucion[tarea.nombre] || 0);
        const inicio = acumulado;
        acumulado += (unidades / total) * 100;
        return `${tarea.color} ${inicio.toFixed(2)}% ${acumulado.toFixed(2)}%`;
    });
    graficaDistribucion.style.background = "conic-gradient(" + segmentos.join(", ") + ")";
    leyendaDistribucion.innerHTML = "";

    configuracion.forEach(function (tarea) {
        const unidades = Number(distribucion[tarea.nombre] || 0);
        const porcentaje = (unidades / total) * 100;
        const fila = document.createElement("div");
        fila.className = "share-legend-row";
        fila.innerHTML = `<i class="share-legend-color" style="color:${tarea.color};background:${tarea.color}"></i><strong>${tarea.nombre}</strong><span>${porcentaje.toFixed(1)}%</span><small>${unidades.toLocaleString("es-UY")} unidades</small>`;
        leyendaDistribucion.appendChild(fila);
    });
}


mesAnterior.addEventListener("click", function () {
    mesVisible -= 1;

    if (mesVisible < 0) {
        mesVisible = 11;
        anioVisible -= 1;
    }

    mostrarCalendario();
});


mesSiguiente.addEventListener("click", function () {
    mesVisible += 1;

    if (mesVisible > 11) {
        mesVisible = 0;
        anioVisible += 1;
    }

    mostrarCalendario();
});


mostrarCalendario();
actualizarResumenActividad();
actualizarDisponibilidadPersonal();
actualizarResumenEmbarques();
actualizarResumenOperativoAnterior();
actualizarTareasHoy();

setInterval(actualizarResumenActividad, 5000);
setInterval(actualizarDisponibilidadPersonal, 30000);
setInterval(actualizarResumenEmbarques, 5000);
setInterval(actualizarResumenOperativoAnterior, 30000);
setInterval(actualizarTareasHoy, 30000);
