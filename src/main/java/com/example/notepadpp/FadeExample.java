package com.example.notepadpp;
import javafx.animation.FadeTransition;
import javafx.animation.ScaleTransition;
import javafx.animation.TranslateTransition;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;
import javafx.util.Duration;

public class FadeExample extends Application {

    @Override
    public void start(Stage primaryStage) {
        Button animatedButton = new Button("Animate Me");

        // Play all animations when the button is clicked
        animatedButton.setOnAction(e -> {
            animateFade(animatedButton);
            animateTranslate(animatedButton);
            animateScale(animatedButton);
        });

        StackPane root = new StackPane(animatedButton);
        Scene scene = new Scene(root, 400, 300);
        primaryStage.setScene(scene);
        primaryStage.setTitle("JavaFX Animation Demo");
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }

    // 👇 Fade animation
    private void animateFade(Button node) {
        FadeTransition ft = new FadeTransition(Duration.millis(800), node);
        ft.setFromValue(0.3);
        ft.setToValue(1.0);
        ft.play();
    }

    // 👇 Slide up/down animation
    private void animateTranslate(Button node) {
        TranslateTransition tt = new TranslateTransition(Duration.millis(600), node);
        tt.setFromY(-20);
        tt.setToY(0);
        tt.play();
    }

    // 👇 Zoom in/out animation
    private void animateScale(Button node) {
        ScaleTransition st = new ScaleTransition(Duration.millis(500), node);
        st.setFromX(0.8);
        st.setFromY(0.8);
        st.setToX(1.0);
        st.setToY(1.0);
        st.play();
    }
}
