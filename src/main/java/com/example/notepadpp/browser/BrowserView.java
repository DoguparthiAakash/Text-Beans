package com.example.notepadpp.browser;

import javafx.scene.layout.BorderPane;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;
import javafx.scene.control.TextField;
import javafx.scene.control.Button;
import javafx.scene.layout.HBox;

public class BrowserView extends BorderPane {

    public BrowserView() {
        WebView webView = new WebView();
        WebEngine webEngine = webView.getEngine();

        TextField urlField = new TextField("https://www.google.com");
        Button goButton = new Button("Go");

        goButton.setOnAction(e -> {
            String url = urlField.getText();
            if (!url.startsWith("http")) url = "http://" + url;
            webEngine.load(url);
        });

        urlField.setOnAction(e -> goButton.fire());

        HBox topBar = new HBox(urlField, goButton);
        setTop(topBar);
        setCenter(webView);

        webEngine.load(urlField.getText());
    }
}
