package com.example.notepadpp.media;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.net.URL;

import javax.swing.ImageIcon;
public class IconUtils {
    
    /**
     * Loads an icon from the classpath with specified dimensions
     * @param path The path to the icon resource (e.g., "/icons/run.png")
     * @param width The desired width of the icon
     * @param height The desired height of the icon
     * @return The scaled ImageIcon, or null if not found
     */
    public static ImageIcon createIcon(String path, int width, int height) {
        try {
            URL resource = IconUtils.class.getResource(path);
            if (resource == null) {
                System.err.println("Icon not found: " + path);
                return createPlaceholderIcon(width, height, Color.RED);
            }
            
            ImageIcon originalIcon = new ImageIcon(resource);
            Image scaledImage = originalIcon.getImage().getScaledInstance(width, height, Image.SCALE_SMOOTH);
            return new ImageIcon(scaledImage);
        } catch (Exception e) {
            System.err.println("Error loading icon: " + path + " - " + e.getMessage());
            return createPlaceholderIcon(width, height, Color.ORANGE);
        }
    }
    
    /**
     * Loads an icon from the classpath with original dimensions
     * @param path The path to the icon resource
     * @return The original ImageIcon, or null if not found
     */
    public static ImageIcon createIcon(String path) {
        try {
            URL resource = IconUtils.class.getResource(path);
            if (resource == null) {
                System.err.println("Icon not found: " + path);
                return null;
            }
            return new ImageIcon(resource);
        } catch (Exception e) {
            System.err.println("Error loading icon: " + path + " - " + e.getMessage());
            return null;
        }
    }
    
    /**
     * Creates a placeholder icon when the real icon is not available
     * @param width The width of the placeholder
     * @param height The height of the placeholder
     * @param color The color of the placeholder
     * @return A generated placeholder icon
     */
    private static ImageIcon createPlaceholderIcon(int width, int height, Color color) {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();
        
        // Set rendering hints for better quality
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        // Draw background
        g2d.setColor(color);
        g2d.fillRect(0, 0, width, height);
        
        // Draw border
        g2d.setColor(color.darker());
        g2d.drawRect(0, 0, width - 1, height - 1);
        
        // Draw "X" mark
        g2d.setColor(Color.WHITE);
        g2d.setStroke(new BasicStroke(2));
        g2d.drawLine(2, 2, width - 2, height - 2);
        g2d.drawLine(width - 2, 2, 2, height - 2);
        
        g2d.dispose();
        return new ImageIcon(image);
    }
    
    /**
     * Creates a colored circle icon (useful for language indicators)
     * @param color The color of the circle
     * @param diameter The diameter of the circle
     * @return A circular icon
     */
    public static ImageIcon createColoredCircleIcon(Color color, int diameter) {
        BufferedImage image = new BufferedImage(diameter, diameter, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();
        
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setColor(color);
        g2d.fillOval(0, 0, diameter, diameter);
        
        g2d.setColor(color.darker());
        g2d.drawOval(0, 0, diameter - 1, diameter - 1);
        
        g2d.dispose();
        return new ImageIcon(image);
    }
    
    /**
     * Creates a monochrome icon from a shape (useful for simple UI icons)
     * @param shape The shape to draw (e.g., triangle for play, square for stop)
     * @param color The color of the icon
     * @param size The size of the icon
     * @return A shape-based icon
     */
    public static ImageIcon createShapeIcon(IconShape shape, Color color, int size) {
        BufferedImage image = new BufferedImage(size, size, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();
        
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setColor(color);
        
        int padding = size / 4;
        int drawSize = size - (2 * padding);
        
        switch (shape) {
            case PLAY:
                int[] xPoints = {padding, padding, size - padding};
                int[] yPoints = {padding, size - padding, size / 2};
                g2d.fillPolygon(xPoints, yPoints, 3);
                break;
                
            case STOP:
                g2d.fillRect(padding, padding, drawSize, drawSize);
                break;
                
            case PAUSE:
                g2d.fillRect(padding, padding, drawSize / 3, drawSize);
                g2d.fillRect(size - padding - (drawSize / 3), padding, drawSize / 3, drawSize);
                break;
                
            case CIRCLE:
                g2d.fillOval(padding, padding, drawSize, drawSize);
                break;
        }
        
        g2d.dispose();
        return new ImageIcon(image);
    }
    
    /**
     * Enum for predefined shapes
     */
    public enum IconShape {
        PLAY, STOP, PAUSE, CIRCLE
    }
}