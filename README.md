# 📧 IMAP Alert Checker - Tabla Automática de Últimas Alertas

Esta herramienta en Python permite consultar vía IMAP la última alerta enviada por correo para cada cliente, según filtros personalizados definidos en un archivo Excel. El script actualiza automáticamente la fecha del último evento en la tabla.

---

## 🧩 Archivos principales

| Archivo               | Descripción                                                  |
|------------------------|--------------------------------------------------------------|
| `actualizar_tabla.py` | Script principal que ejecuta la búsqueda y actualización.     |
| `TB-CL.xlsx`           | Archivo Excel con los criterios por cliente. No se sube a GitHub. |
| `.env`                | Contiene credenciales y configuraciones.                     |

---

## 🧠 Requisitos

- Python 3.8+
- Conexión IMAP habilitada (Gmail o corporativo)
- Acceso vía VPN (si estás en una red monitoreada como CCI)

---

## 🛠️ Instalación

```bash
git clone https://github.com/tuusuario/imap-alert-checker.git
cd imap-alert-checker
pip install -r requirements.txt
