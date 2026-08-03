package com.rodririta.dashboard;

import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.graphics.Color;
import android.os.Build;
import android.os.Bundle;
import android.view.View;
import android.webkit.JavascriptInterface;
import androidx.core.graphics.Insets;
import androidx.core.view.ViewCompat;
import androidx.core.view.WindowInsetsCompat;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        createNotificationChannel();
        getWindow().setStatusBarColor(Color.parseColor("#1e293b"));
        View content = findViewById(android.R.id.content);
        ViewCompat.setOnApplyWindowInsetsListener(content, (v, insets) -> {
            Insets bars = insets.getInsets(WindowInsetsCompat.Type.statusBars());
            v.setPadding(0, bars.top, 0, 0);
            return insets;
        });
        getBridge().getWebView().addJavascriptInterface(new Object() {
            @JavascriptInterface
            public void setStatusBarBgColor(final String hex) {
                runOnUiThread(() -> {
                    try {
                        getWindow().getDecorView().setBackgroundColor(Color.parseColor(hex));
                    } catch (Exception ignored) {}
                });
            }
        }, "NativeApp");
    }

    private void createNotificationChannel() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.O) return;
        NotificationChannel channel = new NotificationChannel(
            "dashboard_high",
            "Dashboard — Vencimientos",
            NotificationManager.IMPORTANCE_HIGH
        );
        channel.setDescription("Recordatorios de fechas límite");
        channel.enableLights(true);
        channel.enableVibration(true);
        NotificationManager nm = getSystemService(NotificationManager.class);
        if (nm != null) nm.createNotificationChannel(channel);
    }
}
