package com.example.notepadpp.editors;

public interface SpecializedEditor {
    void applySyntaxHighlighting();
    void applyFormatting();
    void enableCollaboration();
    void addComment(String comment);
    void showVersionHistory();
}