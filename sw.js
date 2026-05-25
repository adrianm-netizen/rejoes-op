
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(clients.claim()));

// Show notification
self.addEventListener('push', e => {
  const data = e.data ? e.data.json() : {};
  e.waitUntil(self.registration.showNotification(data.title || 'Rejoes', {
    body: data.body || '',
    icon: '/rejoes-op/icon-192.png',
    badge: '/rejoes-op/icon-192.png',
    vibrate: [200, 100, 200],
    tag: 'rejoes-reminder',
    requireInteraction: true,
    actions: [
      { action: 'open', title: '📊 Deschide aplicația' },
      { action: 'dismiss', title: 'Închide' }
    ]
  }));
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  if (e.action === 'open' || !e.action) {
    e.waitUntil(clients.openWindow('/rejoes-op/'));
  }
});
