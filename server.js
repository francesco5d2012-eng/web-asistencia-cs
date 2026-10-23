const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const cors = require('cors');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');

const app = express();
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

app.use(cors());
app.use(express.json());

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

// 1. Obtener lista única de Sedes Jurisdiccionales para el filtro
app.get('/api/sedes', async (req, res) => {
    try {
        const result = await pool.query(
            `SELECT DISTINCT sede_jurisdiccional FROM participantes WHERE sede_jurisdiccional IS NOT NULL ORDER BY sede_jurisdiccional ASC`
        );
        res.json(result.rows.map(r => r.sede_jurisdiccional));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 2. Obtener reporte principal (con filtros de turno, sede, búsqueda y estado)
app.get('/api/reporte', async (req, res) => {
    try {
        const { col = 10, buscar = '', estado = 'TODOS', sede = 'TODAS' } = req.query;

        let sql = `
            SELECT p.orden, p.dni, p.apellido_paterno, p.apellido_materno, p.nombres, 
                   p.sede_regional, p.sede_jurisdiccional,
                   CASE WHEN a.dni IS NOT NULL THEN 'PRESENTE' ELSE 'FALTANTE' END AS estado,
                   to_char(a.fecha_hora, 'YYYY-MM-DD HH24:MI:SS') AS fecha_hora,
                   (SELECT tipo || ' (' || estado || ')' FROM incidencias i WHERE i.dni = p.dni AND i.col_turno = $1 ORDER BY id DESC LIMIT 1) AS incidencia
            FROM participantes p
            LEFT JOIN asistencias a ON p.dni = a.dni AND a.col_turno = $1
            WHERE 1=1
        `;
        let params = [col];

        if (sede !== 'TODAS' && sede.trim() !== '') {
            params.push(sede.trim());
            sql += ` AND p.sede_jurisdiccional = $${params.length}`;
        }

        if (buscar.trim() !== '') {
            params.push(`%${buscar.trim()}%`);
            sql += ` AND (p.dni ILIKE $${params.length} OR p.nombres ILIKE $${params.length} OR p.apellido_paterno ILIKE $${params.length} OR p.apellido_materno ILIKE $${params.length})`;
        }

        sql += ` ORDER BY CAST(p.orden AS INTEGER) ASC`;

        const result = await pool.query(sql, params);
        let asistentes = result.rows;

        const total = asistentes.length;
        const presentes = asistentes.filter(i => i.estado === 'PRESENTE').length;
        const faltantes = total - presentes;

        if (estado === 'PRESENTE') {
            asistentes = asistentes.filter(i => i.estado === 'PRESENTE');
        } else if (estado === 'FALTANTE') {
            asistentes = asistentes.filter(i => i.estado === 'FALTANTE');
        }

        res.json({ total, presentes, faltantes, asistentes });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 3. Registrar asistencia (Desde App Móvil Kotlin) y emitir por WebSocket
app.post('/api/asistencia', async (req, res) => {
    try {
        const { dni, col_turno, etiqueta_turno } = req.body;

        const partCheck = await pool.query(`SELECT * FROM participantes WHERE dni = $1`, [dni]);
        if (partCheck.rows.length === 0) {
            return res.status(404).json({ error: 'DNI no encontrado en el padrón' });
        }

        const insertQuery = `
            INSERT INTO asistencias (dni, col_turno, etiqueta_turno)
            VALUES ($1, $2, $3)
            ON CONFLICT (dni, col_turno) 
            DO UPDATE SET fecha_hora = (CURRENT_TIMESTAMP AT TIME ZONE 'America/Lima')
            RETURNING *;
        `;
        const result = await pool.query(insertQuery, [dni, col_turno, etiqueta_turno || `Turno ${col_turno}`]);

        // Notificar a todos los paneles web abiertos en tiempo real
        io.emit('actualizacion_asistencia', {
            dni,
            col_turno,
            participante: partCheck.rows[0]
        });

        res.json({ ok: true, registro: result.rows[0], participante: partCheck.rows[0] });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 4. Descargar Reporte Consolidado en Excel (.xlsx)
app.get('/api/reportes/excel', async (req, res) => {
    try {
        const { col = 10, sede = 'TODAS', estado = 'TODOS', buscar = '' } = req.query;

        let sql = `
            SELECT p.orden, p.dni, p.apellido_paterno, p.apellido_materno, p.nombres, 
                   p.sede_regional, p.sede_jurisdiccional,
                   CASE WHEN a.dni IS NOT NULL THEN 'PRESENTE' ELSE 'FALTANTE' END AS estado,
                   to_char(a.fecha_hora, 'YYYY-MM-DD HH24:MI:SS') AS fecha_hora,
                   (SELECT tipo || ' - ' || estado || ': ' || COALESCE(observaciones,'') FROM incidencias i WHERE i.dni = p.dni AND i.col_turno = $1 ORDER BY id DESC LIMIT 1) AS incidencia
            FROM participantes p
            LEFT JOIN asistencias a ON p.dni = a.dni AND a.col_turno = $1
            WHERE 1=1
        `;
        let params = [col];

        if (sede !== 'TODAS' && sede.trim() !== '') {
            params.push(sede.trim());
            sql += ` AND p.sede_jurisdiccional = $${params.length}`;
        }

        if (buscar.trim() !== '') {
            params.push(`%${buscar.trim()}%`);
            sql += ` AND (p.dni ILIKE $${params.length} OR p.nombres ILIKE $${params.length} OR p.apellido_paterno ILIKE $${params.length})`;
        }

        sql += ` ORDER BY CAST(p.orden AS INTEGER) ASC`;

        const result = await pool.query(sql, params);
        let filas = result.rows;

        if (estado !== 'TODOS') {
            filas = filas.filter(r => r.estado === estado);
        }

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Asistencia CS 2026');

        sheet.columns = [
            { header: 'Nº', key: 'orden', width: 6 },
            { header: 'DNI', key: 'dni', width: 12 },
            { header: 'Apellido Paterno', key: 'apellido_paterno', width: 18 },
            { header: 'Apellido Materno', key: 'apellido_materno', width: 18 },
            { header: 'Nombres', key: 'nombres', width: 25 },
            { header: 'Sede Regional', key: 'sede_regional', width: 20 },
            { header: 'Sede Jurisdiccional', key: 'sede_jurisdiccional', width: 24 },
            { header: 'Estado', key: 'estado', width: 14 },
            { header: 'Hora Registro', key: 'fecha_hora', width: 20 },
            { header: 'Incidencia / Justificación', key: 'incidencia', width: 35 }
        ];

        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D47A1' } };

        sheet.addRows(filas);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Reporte_Asistencia_Turno_${col}.xlsx`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 5. Registrar y actualizar Incidencias
app.post('/api/incidencias', async (req, res) => {
    try {
        const { dni, col_turno, tipo, estado = 'Pendiente', observaciones } = req.body;
        const result = await pool.query(
            `INSERT INTO incidencias (dni, col_turno, tipo, estado, observaciones) VALUES ($1, $2, $3, $4, $5) RETURNING *`,
            [dni, col_turno, tipo, estado, observaciones]
        );
        io.emit('actualizacion_asistencia', { dni, col_turno });
        res.json({ ok: true, incidencia: result.rows[0] });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

io.on('connection', (socket) => {
    console.log('Panel Web conectado en tiempo real:', socket.id);
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`Servidor corriendo en puerto ${PORT}`));