/* Service worker pentru notificări push FCM (Rejoes) — mesaje primite când aplicația e închisă */
importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyBEKvd_Fw0eoLYtLQENUpAqJLKOxZQr6kY",
  authDomain: "rejoes-op.firebaseapp.com",
  projectId: "rejoes-op",
  messagingSenderId: "633578390248",
  appId: "1:633578390248:web:e9e65434d5ae6c84b5ff5d"
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage(function(payload) {
  const n = (payload && payload.notification) || {};
  const title = n.title || '🚨 Rejoes — Alertă pontaj';
  self.registration.showNotification(title, {
    body: n.body || '',
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    tag: 'rejoes-pontaj',
    renotify: true
  });
});

// La click pe notificare, deschide/aduce în față aplicația
self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(list) {
      for (const c of list) { if ('focus' in c) return c.focus(); }
      if (clients.openWindow) return clients.openWindow('Rejoes_Control.html');
    })
  );
});
