"""Capa de compatibilidad para SQLite local y PostgreSQL en la nube.

La aplicación continúa usando SQLite cuando se ejecuta en la PC. Al definir
DATABASE_URL con una URL postgres:// o postgresql://, usa PostgreSQL sin
cambiar los módulos funcionales ni requerir una instalación local adicional.
"""
from __future__ import annotations

import os
import re
import sqlite3


class DatabaseError(Exception):
    """Error de base de datos independiente del proveedor."""


class IntegrityError(DatabaseError):
    """Violación de una clave única o relación requerida."""


class CompatRow(dict):
    """Fila compatible con sqlite3.Row: permite clave o índice."""

    def __getitem__(self, clave):
        if isinstance(clave, int):
            return list(self.values())[clave]
        return super().__getitem__(clave)


class CursorPostgres:
    def __init__(self, cursor, lastrowid=None):
        self._cursor = cursor
        self.lastrowid = lastrowid

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def fetchone(self):
        fila = self._cursor.fetchone()
        return CompatRow(fila) if fila is not None else None

    def fetchall(self):
        return [CompatRow(fila) for fila in self._cursor.fetchall()]

    def __iter__(self):
        for fila in self._cursor:
            yield CompatRow(fila)


class CursorMemoria:
    """Cursor mínimo para emular PRAGMA table_info en PostgreSQL."""

    def __init__(self, filas):
        self._filas = [CompatRow(fila) for fila in filas]
        self.rowcount = len(self._filas)
        self.lastrowid = None

    def fetchone(self):
        return self._filas[0] if self._filas else None

    def fetchall(self):
        return self._filas[:]


TABLAS_CON_ID = {
    "usuarios", "tareas_supervision", "archivos_tareas_supervision",
    "planificaciones_horarios", "turnos_personal", "ausencias_funcionarios",
    "colectores", "sesiones_handheld", "actividades", "periodos_actividad",
    "embarques", "articulos_embarque", "controles_embarque",
    "movimientos_control_embarque", "cargas_metricas", "movimientos_metricas"
}


def es_postgres():
    url = os.environ.get("DATABASE_URL", "").strip().lower()
    return url.startswith("postgres://") or url.startswith("postgresql://")


def _adaptar_sql(sql):
    """Convierte el subconjunto SQLite usado por el proyecto a PostgreSQL."""
    consulta = sql.strip()
    consulta = re.sub(
        r"id\s+INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT",
        "id BIGSERIAL PRIMARY KEY", consulta, flags=re.IGNORECASE
    )
    consulta = re.sub(r"\bAUTOINCREMENT\b", "", consulta, flags=re.IGNORECASE)
    consulta = re.sub(r"\bBLOB\b", "BYTEA", consulta, flags=re.IGNORECASE)
    consulta = re.sub(r"\bINSERT\s+OR\s+IGNORE\s+INTO\b", "INSERT INTO", consulta, flags=re.IGNORECASE)
    consulta = re.sub(
        r"GROUP_CONCAT\(\s*DISTINCT\s+([^)]+?)\s*\)",
        r"STRING_AGG(DISTINCT \1, ',')", consulta, flags=re.IGNORECASE
    )
    consulta = re.sub(
        r"GROUP_CONCAT\((.+?),\s*([^)]+)\)",
        r"STRING_AGG(\1, \2)", consulta, flags=re.IGNORECASE
    )
    consulta = consulta.replace("?", "%s")
    consulta = re.sub(r"\bLIKE\s+(%s)\s+COLLATE\s+NOCASE", r"ILIKE \1", consulta, flags=re.IGNORECASE)
    consulta = re.sub(r"\s+COLLATE\s+NOCASE", "", consulta, flags=re.IGNORECASE)
    return consulta


def _tabla_insertada(sql):
    coincidencia = re.search(r"^\s*INSERT\s+INTO\s+([a-zA-Z_][a-zA-Z0-9_]*)", sql, re.IGNORECASE)
    return coincidencia.group(1).lower() if coincidencia else ""


class ConexionPostgres:
    def __init__(self, conexion):
        self._conexion = conexion

    def _columnas_tabla(self, tabla):
        cursor = self._conexion.cursor(row_factory=self._dict_row)
        cursor.execute("""
            SELECT column_name AS name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = %s
            ORDER BY ordinal_position
        """, (tabla,))
        return CursorMemoria(cursor.fetchall())

    @property
    def _dict_row(self):
        from psycopg.rows import dict_row
        return dict_row

    def execute(self, sql, parametros=()):
        pragma = re.match(r"^\s*PRAGMA\s+table_info\(([^)]+)\)", sql, re.IGNORECASE)
        if pragma:
            return self._columnas_tabla(pragma.group(1).strip(" '`\""))

        consulta = _adaptar_sql(sql)
        tabla = _tabla_insertada(consulta)
        necesita_id = tabla in TABLAS_CON_ID and "returning" not in consulta.lower()
        if "INSERT OR IGNORE" in sql.upper() and "ON CONFLICT" not in consulta.upper():
            consulta = consulta.rstrip(";") + " ON CONFLICT DO NOTHING"
        if necesita_id:
            consulta = consulta.rstrip(";") + " RETURNING id"

        try:
            cursor = self._conexion.cursor(row_factory=self._dict_row)
            cursor.execute(consulta, parametros or ())
            ultimo_id = None
            if necesita_id:
                fila = cursor.fetchone()
                ultimo_id = fila["id"] if fila else None
            return CursorPostgres(cursor, ultimo_id)
        except Exception as error:
            self._manejar_error(error)

    def executemany(self, sql, secuencia):
        ultimo = None
        for parametros in secuencia:
            ultimo = self.execute(sql, parametros)
        if ultimo is None:
            return CursorMemoria([])
        return ultimo

    def _manejar_error(self, error):
        try:
            from psycopg import IntegrityError as PsycopgIntegrityError
            if isinstance(error, PsycopgIntegrityError):
                raise IntegrityError(str(error)) from error
        except ImportError:
            pass
        raise DatabaseError(str(error)) from error

    def commit(self):
        self._conexion.commit()

    def rollback(self):
        self._conexion.rollback()

    def close(self):
        self._conexion.close()

    def __enter__(self):
        return self

    def __exit__(self, tipo_error, valor_error, traza):
        if tipo_error is None:
            self.commit()
        else:
            self.rollback()
        return False


def conectar(ruta_sqlite):
    """Abre PostgreSQL cloud cuando DATABASE_URL está definida; si no, SQLite."""
    if not es_postgres():
        conexion = sqlite3.connect(ruta_sqlite, timeout=10)
        conexion.row_factory = sqlite3.Row
        return conexion

    try:
        import psycopg
        return ConexionPostgres(psycopg.connect(os.environ["DATABASE_URL"]))
    except ImportError as error:
        raise RuntimeError("Falta psycopg. Instala las dependencias del proyecto.") from error
