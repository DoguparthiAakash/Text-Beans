package com.example.notepadpp.media;
import javax.swing.ImageIcon;
public class LanguageItem {
    private String name;
    private ImageIcon icon;
    
    public LanguageItem(String name, ImageIcon icon) {
        this.name = name;
        this.icon = icon;
    }
    
    public LanguageItem(String name, String iconPath) {
        this.name = name;
        this.icon = IconUtils.createIcon(iconPath, 16, 16);
    }
    
    public String getName() { return name; }
    public ImageIcon getIcon() { return icon; }
    
    @Override
    public String toString() {
        return name;
    }
}