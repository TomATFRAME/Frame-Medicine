// Native app initialization — loaded by Capacitor shell
// Bridges web app with native device features (push, haptics, status bar)

document.addEventListener('DOMContentLoaded', function() {
  if (!window.Capacitor) return; // Only run in native shell

  var Capacitor = window.Capacitor;
  var Plugins = Capacitor.Plugins;

  // ---- Status Bar ----
  if (Plugins.StatusBar) {
    Plugins.StatusBar.setStyle({ style: 'DARK' });
    Plugins.StatusBar.setBackgroundColor({ color: '#080808' });
  }

  // ---- Push Notifications ----
  if (Plugins.PushNotifications) {
    var Push = Plugins.PushNotifications;

    // Request permission on first launch
    Push.requestPermissions().then(function(result) {
      if (result.receive === 'granted') {
        Push.register();
      }
    });

    // Get device token — send to backend for storage
    Push.addListener('registration', function(token) {
      console.log('Push token:', token.value);
      // Store token for sending to backend after login
      window._nativePushToken = token.value;
    });

    Push.addListener('registrationError', function(err) {
      console.log('Push registration error:', err);
    });

    // Handle push received while app is open
    Push.addListener('pushNotificationReceived', function(notification) {
      // Show in-app toast if available
      if (typeof showToast === 'function') {
        showToast(notification.title + ': ' + notification.body);
      }
    });

    // Handle push tap — deep link to relevant screen
    Push.addListener('pushNotificationActionPerformed', function(notification) {
      var data = notification.notification.data || {};
      if (data.screen) {
        if (typeof showScreen === 'function') {
          showScreen(data.screen);
        }
      }
    });
  }

  // ---- Haptic Feedback ----
  if (Plugins.Haptics) {
    // Add haptic feedback to buttons
    document.addEventListener('click', function(e) {
      if (e.target.classList.contains('btn') || e.target.classList.contains('nav-item')) {
        Plugins.Haptics.impact({ style: 'LIGHT' });
      }
    });
  }

  // ---- Hide Splash Screen after app loads ----
  if (Plugins.SplashScreen) {
    setTimeout(function() {
      Plugins.SplashScreen.hide();
    }, 500);
  }
});
