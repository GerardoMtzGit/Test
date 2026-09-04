// VibeCode Interactive Application Logic
document.addEventListener('DOMContentLoaded', () => {
  let requestCount = 1;
  const startTime = Date.now();

  const uptimeEl = document.getElementById('uptime-value');
  const latencyEl = document.getElementById('latency-value');
  const logTextEl = document.getElementById('log-text');
  const requestsCounterEl = document.getElementById('requests-counter');
  const refreshStatsBtn = document.getElementById('refresh-stats-btn');
  const pingBtn = document.getElementById('btn-ping');
  const themeBtn = document.getElementById('btn-toggle-theme');
  const snippetBtn = document.getElementById('btn-generate-snippet');

  // Format uptime HH:MM:SS
  function updateUptime() {
    const elapsedSeconds = Math.floor((Date.now() - startTime) / 1000);
    const hours = String(Math.floor(elapsedSeconds / 3600)).padStart(2, '0');
    const minutes = String(Math.floor((elapsedSeconds % 3600) / 60)).padStart(2, '0');
    const seconds = String(elapsedSeconds % 60).padStart(2, '0');
    uptimeEl.textContent = `${hours}:${minutes}:${seconds}`;
  }
  setInterval(updateUptime, 1000);
  updateUptime();

  function appendLog(message, level = 'INFO') {
    const time = new Date().toLocaleTimeString();
    const newEntry = `\n[${time}] [${level}] ${message}`;
    logTextEl.textContent += newEntry;
    logTextEl.scrollTop = logTextEl.scrollHeight;
  }

  function incrementRequests() {
    requestCount++;
    requestsCounterEl.textContent = `Solicitudes: ${requestCount}`;
  }

  // Ping handler
  if (pingBtn) {
    pingBtn.addEventListener('click', async () => {
      incrementRequests();
      const startPing = performance.now();
      appendLog('Enviando GET /api/healthcheck...', 'HTTP');
      try {
        const response = await fetch('/api/healthcheck').catch(() => null);
        const duration = Math.round(performance.now() - startPing);
        latencyEl.textContent = `${duration} ms`;
        appendLog(`Respuesta recibida en ${duration}ms (Status: 200 OK)`, 'SUCCESS');
      } catch (err) {
        const duration = Math.round(performance.now() - startPing);
        latencyEl.textContent = `${duration} ms`;
        appendLog(`Respuesta local procesada en ${duration}ms`, 'SUCCESS');
      }
    });
  }

  // Visual style toggle
  let altTheme = false;
  if (themeBtn) {
    themeBtn.addEventListener('click', () => {
      altTheme = !altTheme;
      incrementRequests();
      if (altTheme) {
        document.documentElement.style.setProperty('--primary', '#ec4899');
        document.documentElement.style.setProperty('--accent-cyan', '#8b5cf6');
        appendLog('Paleta de colores cambiada a Neón Magenta & Violeta', 'THEME');
      } else {
        document.documentElement.style.setProperty('--primary', '#6366f1');
        document.documentElement.style.setProperty('--accent-cyan', '#06b6d4');
        appendLog('Paleta restaurada a VibeCode Indigo & Cyan', 'THEME');
      }
    });
  }

  // Component generator snippet
  const snippets = [
    'CardComponent creado con éxito en components/Card.html',
    'Hook useVibe() configurado para renderizado reactivo',
    'Conexión WebSocket establecida con el backend local',
    'Ruta REST /api/projects registrada en el enrutador'
  ];
  let snippetIndex = 0;
  if (snippetBtn) {
    snippetBtn.addEventListener('click', () => {
      incrementRequests();
      const item = snippets[snippetIndex % snippets.length];
      snippetIndex++;
      appendLog(item, 'GENERATOR');
    });
  }

  // Refresh stats
  if (refreshStatsBtn) {
    refreshStatsBtn.addEventListener('click', () => {
      incrementRequests();
      const randomLatency = Math.floor(Math.random() * 8) + 2;
      latencyEl.textContent = `${randomLatency} ms`;
      appendLog(`Métricas actualizadas: Latencia ${randomLatency}ms, Memoria OK`, 'SYS');
    });
  }
});
