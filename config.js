// Configuración de Hub Evaluación 360.
// Lo que el administrador guarde en la pantalla "Configuración" de la app manda sobre los valores iniciales de aquí.
window.HUB_CONFIG = {
  // 'demo'      : todo queda en este navegador y se elige a quién "ser" desde la barra superior (también con …/#demo).
  // 'sharepoint': ingreso con la cuenta Microsoft de la empresa y registros en una lista de SharePoint (README.md).
  modo: 'sharepoint',

  // Lista de SharePoint donde se guarda todo (una fila por respuesta + una fila de configuración).
  // Sitio propio y restringido: https://marcalman.sharepoint.com/sites/Evaluacion360 (ver README, "Confidencialidad").
  sharepoint: {
    sitio: 'marcalman.sharepoint.com:/sites/Evaluacion360',
    lista: 'Evaluacion360'
  },

  // Registro de la app en Microsoft Entra ID: se reutiliza el de "Hub Comex" (permiso AllSites.Write aprobado el 2026-10-27).
  // Cada dirección donde se abra la app debe estar en Autenticación → SPA (hoy: http://localhost:5500/ para pruebas).
  // Mientras esté vacío, la app solo funciona en modo demo. La demo sigue disponible con …/?demo
  clientId: '3abc78ee-335d-4231-a7d2-4dab471900a5',
  tenantId: '8f9b210d-f5e5-404f-9fed-a0a827154105',
  permisos: ['https://marcalman.sharepoint.com/AllSites.Write'],

  // Siempre pueden entrar a Configuración y al Panel de Talento Humano.
  administradores: ['sistemasdegestion@healthyfood.com.ec', 'talentohumano@healthyfood.com.ec'],

  // Personas, jefes y correos: viven en SharePoint (Configuración → Personas). No se ponen aquí porque este archivo es público.
  // La copia local los toma de datos-personas.js (no publicado) para la demo y para la primera carga.
  personas: []
};
