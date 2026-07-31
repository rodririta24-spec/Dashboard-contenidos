const { onSchedule } = require('firebase-functions/v2/scheduler');
const admin = require('firebase-admin');

admin.initializeApp();

// Runs every day at 9:00 AM Argentina time
exports.checkDeadlines = onSchedule(
  { schedule: '0 9 * * *', timeZone: 'America/Argentina/Buenos_Aires' },
  async () => {
    const db = admin.firestore();

    // Get the Apps Script API URL from Firestore config
    const configSnap = await db.collection('config').doc('settings').get();
    const apiUrl = configSnap.data()?.apiUrl;
    if (!apiUrl) {
      console.log('No API URL configured in Firestore config/settings');
      return;
    }

    // Fetch project data from Apps Script
    let projects = [];
    try {
      const res = await fetch(apiUrl);
      const json = await res.json();
      projects = json.data || [];
    } catch (err) {
      console.error('Failed to fetch projects:', err);
      return;
    }

    // Find projects with deadline tomorrow (not yet delivered)
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);

    function parseSheetDate(s) {
      if (!s) return null;
      const parts = String(s).trim().split('/');
      if (parts.length !== 3) return null;
      const [d, m, y] = parts.map(Number);
      return new Date(y, m - 1, d);
    }

    function isTomorrow(dateStr) {
      const d = parseSheetDate(dateStr);
      if (!d) return false;
      return (
        d.getFullYear() === tomorrow.getFullYear() &&
        d.getMonth()    === tomorrow.getMonth()    &&
        d.getDate()     === tomorrow.getDate()
      );
    }

    const dueTomorrow = projects.filter(p =>
      isTomorrow(p['Fecha Limite']) &&
      !String(p['Estado'] || '').toLowerCase().includes('entregado')
    );

    if (dueTomorrow.length === 0) {
      console.log('No deadlines tomorrow — no notifications sent');
      return;
    }

    // Build notification text
    const names = dueTomorrow.map(p => p['Titulo']).slice(0, 3).join(', ');
    const extra = dueTomorrow.length > 3 ? ` y ${dueTomorrow.length - 3} más` : '';
    const body  = dueTomorrow.length === 1
      ? `"${dueTomorrow[0]['Titulo']}" vence mañana`
      : `${dueTomorrow.length} proyectos vencen mañana: ${names}${extra}`;

    // Get all registered FCM tokens
    const tokensSnap = await db.collection('fcmTokens').get();
    const tokens = tokensSnap.docs.map(d => d.data().token).filter(Boolean);
    if (tokens.length === 0) {
      console.log('No FCM tokens registered');
      return;
    }

    // Send in batches of 500 (FCM limit)
    const messaging = admin.messaging();
    for (let i = 0; i < tokens.length; i += 500) {
      const chunk = tokens.slice(i, i + 500);
      const result = await messaging.sendEachForMulticast({
        tokens: chunk,
        notification: { title: '📅 Deadlines mañana', body },
        android: {
          priority: 'high',
          notification: { channelId: 'deadlines', sound: 'default' }
        }
      });
      console.log(`Sent ${result.successCount}/${chunk.length} notifications`);

      // Clean up invalid tokens
      const toDelete = [];
      result.responses.forEach((r, idx) => {
        if (!r.success && r.error?.code === 'messaging/registration-token-not-registered') {
          toDelete.push(chunk[idx]);
        }
      });
      for (const token of toDelete) {
        await db.collection('fcmTokens').doc(token).delete().catch(() => {});
      }
    }

    console.log(`Deadline check complete: ${dueTomorrow.length} project(s) due tomorrow`);
  }
);
