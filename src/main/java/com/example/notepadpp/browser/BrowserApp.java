package com.example.notepadpp.browser;

import java.util.ArrayList;
import java.util.List;

import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.concurrent.Worker;
import javafx.embed.swing.JFXPanel;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebHistory;
import javafx.scene.web.WebView;
import javafx.stage.Stage;

public class BrowserApp {
    private static boolean javafxInitialized = false;
    private static final String HOME_URL = "https://www.google.com";
    private static final List<Stage> openStages = new ArrayList<>();

    public static void launchBrowser() {
        ensureJavaFXInitialized();
        Platform.runLater(() -> {
            Stage stage = new Stage();
            WebView webView = new WebView();
            
            Scene scene = createBrowserScene(stage, webView);
            stage.setScene(scene);
            
            stage.getIcons().add(loadIcon("/icons/browser.png"));
            openStages.add(stage);
            stage.show();
            
            webView.getEngine().load(HOME_URL);
        });
    }

    public static void launchMarkdownRenderer(String markdownText) {
        ensureJavaFXInitialized();
        Platform.runLater(() -> {
            Parser parser = Parser.builder().build();
            HtmlRenderer renderer = HtmlRenderer.builder().build();
            Node document = parser.parse(markdownText);
            String html = renderer.render(document);

            WebView webView = new WebView();
            WebEngine engine = webView.getEngine();
            engine.loadContent(html, "text/html");

            Scene scene = new Scene(webView, 800, 600);
            Stage stage = new Stage();
            stage.setTitle("Markdown Preview");
            stage.getIcons().add(loadIcon("/icons/bookmarkline.png"));
            stage.setScene(scene);
            
            // STABILITY FIX: Add cleanup logic to the Markdown preview
            stage.setOnCloseRequest(e -> engine.load("about:blank"));
            
            stage.show();
        });
    }
private static void applyUITheme(Scene scene, boolean dark) {
    String cssPath = dark ? "/css/dark-theme.css" : "/css/light-theme.css";
    try {
        String css = BrowserApp.class.getResource(cssPath).toExternalForm();
        scene.getStylesheets().setAll(css);
    } catch (Exception e) {
        System.err.println("Could not load theme CSS: " + cssPath);
    }
}
    private static Scene createBrowserScene(Stage stage, WebView webView) {
        WebEngine engine = webView.getEngine();
        WebHistory history = engine.getHistory();

        // --- Toolbar Components ---
        Button backBtn = createIconButton(loadIcon("/icons/back.png"), "Go back");
        backBtn.setOnAction(e -> history.go(-1));

        Button forwardBtn = createIconButton(loadIcon("/icons/next.png"), "Go forward");
        forwardBtn.setOnAction(e -> history.go(1));
        
        // STABILITY FEATURE: Add a Stop button
        Button stopBtn = createIconButton(loadIcon("/icons/stop.png"), "Stop loading");
        stopBtn.setOnAction(e -> engine.getLoadWorker().cancel());

        Button reloadBtn = createIconButton(loadIcon("/icons/replace.png"), "Reload page");
        reloadBtn.setOnAction(e -> engine.reload());

        Button homeBtn = createIconButton(loadIcon("/icons/home.png"), "Go home");
        homeBtn.setOnAction(e -> engine.load(HOME_URL));
        
        backBtn.disableProperty().bind(history.currentIndexProperty().isEqualTo(0));
        forwardBtn.disableProperty().bind(Bindings.createBooleanBinding(() -> 
            history.getCurrentIndex() >= history.getEntries().size() - 1, history.currentIndexProperty(), history.getEntries()));
        
        TextField addressBar = new TextField();
        HBox.setHgrow(addressBar, Priority.ALWAYS);
        addressBar.setOnAction(e -> {
            String url = addressBar.getText();
            if (!url.startsWith("http")) url = "http://" + url;
            engine.load(url);
        });

        // STABILITY FEATURE: Restore Dark Mode for web content
        ToggleButton darkModeToggle = new ToggleButton("Dark Mode");
        String darkWebCss = BrowserApp.class.getResource("/css/dark-web.css").toExternalForm();
        darkModeToggle.setOnAction(e -> {
            if (darkModeToggle.isSelected()) {
                engine.setUserStyleSheetLocation(darkWebCss);
            } else {
                engine.setUserStyleSheetLocation(null);
            }
        });
        
        ProgressBar progressBar = new ProgressBar();
        progressBar.progressProperty().bind(engine.getLoadWorker().progressProperty());
        progressBar.visibleProperty().bind(engine.getLoadWorker().stateProperty().isEqualTo(Worker.State.RUNNING));
        
        // Only show Stop button when loading
        stopBtn.visibleProperty().bind(progressBar.visibleProperty());

        // --- Listeners for Stability and Features ---
        engine.getLoadWorker().stateProperty().addListener((obs, oldState, newState) -> {
            if (newState == Worker.State.SUCCEEDED) {
                // Update window title with page title
                stage.setTitle(engine.getTitle());
                // Re-apply dark mode CSS if enabled
                if (darkModeToggle.isSelected()) {
                    engine.setUserStyleSheetLocation(darkWebCss);
                }
            }
        });
        engine.locationProperty().addListener((obs, oldVal, newVal) -> addressBar.setText(newVal));

        // STABILITY FIX: Add cleanup logic when the main browser window is closed
        stage.setOnCloseRequest(e -> {
            engine.getLoadWorker().cancel();
            engine.load("about:blank");
            openStages.remove(stage);
        });
        
        // --- Layout ---
        HBox toolbar = new HBox(5, backBtn, forwardBtn, stopBtn, reloadBtn, homeBtn, addressBar, darkModeToggle);
        toolbar.setPadding(new Insets(5));
        
        BorderPane root = new BorderPane();
        root.setTop(new BorderPane(toolbar, progressBar, null, null, null));
        root.setCenter(webView);

        return new Scene(root, 1200, 800);
    }

    // --- Helper Methods ---

    private static void ensureJavaFXInitialized() {
        if (!javafxInitialized) {
            new JFXPanel(); // This initializes the JavaFX toolkit.
            Platform.setImplicitExit(false);
            javafxInitialized = true;
        }
    }

    private static Button createIconButton(Image icon, String tooltip) {
        Button btn = new Button();
        if (icon != null) btn.setGraphic(new ImageView(icon));
        btn.setTooltip(new Tooltip(tooltip));
        return btn;
    }

    private static Image loadIcon(String path) {
        try {
            return new Image(BrowserApp.class.getResource(path).toExternalForm(), 16, 16, true, true);
        } catch (Exception e) {
            System.err.println("Icon not found: " + path);
            return null;
        }
    }
}