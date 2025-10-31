package com.example.notepadpp.ai;

import ai.djl.Application;
import ai.djl.ModelException;
import ai.djl.inference.Predictor;
import ai.djl.modality.nlp.qa.QAInput;

import ai.djl.translate.TranslateException;
import ai.djl.repository.zoo.Criteria;
import ai.djl.repository.zoo.ZooModel;
import ai.djl.repository.zoo.ZooProvider;
import ai.djl.repository.zoo.ModelZoo;
import ai.djl.training.util.ProgressBar;

import java.io.IOException;

public class LocalDJLClient {

    private static Predictor<QAInput, String> predictor;

    static {
        try {
            Criteria<QAInput, String> criteria = Criteria.builder()
                    .optApplication(Application.NLP.QUESTION_ANSWER)
                    .setTypes(QAInput.class, String.class)
                    .optEngine("PyTorch")
                    .optModelUrls("https://resources.djl.ai/test-models/distilbert_squad-distilled.zip")
                    .optProgress(new ProgressBar())
                    .build();

            ZooModel<QAInput, String> model = criteria.loadModel();
            predictor = model.newPredictor();
        } catch (IOException | ModelException e) {
            e.printStackTrace();
        }
    }

    public static String askAI(String question) {
        if (predictor == null) return "⚠️ AI not initialized.";

        // Provide context manually or dynamically
        String context = """
                Java is a high-level, class-based, object-oriented programming language 
                that is designed to have as few implementation dependencies as possible.
                """;

        QAInput input = new QAInput(question, context);
        try {
            return predictor.predict(input);
        } catch (TranslateException e) {
            return "⚠️ AI error: " + e.getMessage();
        }
    }
}
