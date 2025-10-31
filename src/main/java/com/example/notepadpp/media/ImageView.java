package com.example.notepadpp.media;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.lang.ref.SoftReference;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.swing.AbstractAction;
import javax.swing.ActionMap;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.InputMap;
import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;

/**
 * Memory-efficient ImageView with aggressive memory management
 * Uses SoftReferences, lazy loading, and smart caching to minimize RAM usage
 */
public class ImageView {
    
    // ============================================================================
    // MEMORY-EFFICIENT CACHE - Uses SoftReferences to allow GC when memory is low
    // ============================================================================
    
    // Small icons cache (never exceeds 20 entries, ~100KB total)
    private static final Map<String, Icon> ICON_CACHE = new LinkedHashMap<String, Icon>(16, 0.75f, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry<String, Icon> eldest) {
            return size() > 20; // Limit to 20 small icons max
        }
    };
    
    // Large image cache using SoftReferences (GC can clear when memory is low)
    private static final Map<String, SoftReference<BufferedImage>> IMAGE_CACHE = new LinkedHashMap<>(8, 0.75f, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry<String, SoftReference<BufferedImage>> eldest) {
            return size() > 5; // Keep max 5 large images
        }
    };
    
    // Configuration
    private static final String[] ICON_SEARCH_PATHS = {
        "/icons/", "/images/", "/resources/icons/", "/resources/images/", "/"
    };
    
    // Maximum dimensions for loaded images to prevent memory explosion
    private static final int MAX_IMAGE_DIMENSION = 4096; // 4K max
    private static final int THUMBNAIL_SIZE = 256; // For cache previews
    
    // ============================================================================
    // PUBLIC ICON GETTERS - Minimal memory footprint
    // ============================================================================
    
    public static Icon getTabIcon() {
        return getCachedIcon("imageview.png", 16, 16);
    }
    
    public static Icon getTabIcon(int width, int height) {
        return getCachedIcon("imageview.png", width, height);
    }
    
    public static Icon getCloseIcon() {
        return getCachedIconWithFallback("closetab.png", 12, 12, ImageView::createDefaultCloseIcon);
    }
    
    public static Icon getCloseIcon(int size) {
        return getCachedIconWithFallback("closetab.png", size, size, ImageView::createDefaultCloseIcon);
    }
    
    public static Icon getCloseIcon(int width, int height) {
        return getCachedIconWithFallback("closetab.png", width, height, ImageView::createDefaultCloseIcon);
    }
    
    public static Icon getNewTabIcon() {
        return getCachedIconWithFallback("newtab.png", 16, 16, ImageView::createDefaultNewTabIcon);
    }
    
    public static Icon getNewTabIcon(int width, int height) {
        return getCachedIconWithFallback("newtab.png", width, height, ImageView::createDefaultNewTabIcon);
    }
    
    public static Icon getNewTableTabIcon() {
        return getCachedIconWithFallback("newtable.png", 16, 16, ImageView::createDefaultTableIcon);
    }
    
    public static Icon getNewTableTabIcon(int width, int height) {
        return getCachedIconWithFallback("newtable.png", width, height, ImageView::createDefaultTableIcon);
    }
    
    public static Icon getSaveIcon() {
        return getCachedIcon("save.png", 16, 16);
    }
    
    public static Icon getSaveIcon(int width, int height) {
        return getCachedIcon("save.png", width, height);
    }
    
    public static Icon getOpenIcon() {
        return getCachedIcon("open.png", 16, 16);
    }
    
    public static Icon getOpenIcon(int width, int height) {
        return getCachedIcon("open.png", width, height);
    }
    
    // ============================================================================
    // MEMORY-EFFICIENT CACHING - Only keeps small icons in memory
    // ============================================================================
    
    private static Icon getCachedIcon(String filename, int width, int height) {
        // Only cache small icons (< 32px)
        if (width > 32 || height > 32) {
            return loadIconOptimized(filename, width, height);
        }
        
        String cacheKey = filename + "_" + width + "x" + height;
        Icon cached = ICON_CACHE.get(cacheKey);
        
        if (cached != null) {
            return cached;
        }
        
        Icon icon = loadIconOptimized(filename, width, height);
        if (icon != null) {
            ICON_CACHE.put(cacheKey, icon);
        }
        
        return icon;
    }
    
    private static Icon getCachedIconWithFallback(String filename, int width, int height, 
                                                  IconGenerator fallbackGenerator) {
        // Only cache small icons
        if (width > 32 || height > 32) {
            Icon icon = loadIconOptimized(filename, width, height);
            return (icon != null) ? icon : fallbackGenerator.generate(width, height);
        }
        
        String cacheKey = filename + "_" + width + "x" + height;
        Icon cached = ICON_CACHE.get(cacheKey);
        
        if (cached != null) {
            return cached;
        }
        
        Icon icon = loadIconOptimized(filename, width, height);
        if (icon == null || (icon instanceof ImageIcon && 
            ((ImageIcon)icon).getImage().getWidth(null) <= 1)) {
            icon = fallbackGenerator.generate(width, height);
        }
        
        if (icon != null) {
            ICON_CACHE.put(cacheKey, icon);
        }
        
        return icon;
    }
    
    /**
     * Lightweight icon loading - no caching for large images
     */
    private static Icon loadIconOptimized(String filename, int width, int height) {
        try {
            for (String basePath : ICON_SEARCH_PATHS) {
                String resourcePath = basePath + filename;
                java.net.URL resource = ImageView.class.getResource(resourcePath);
                
                if (resource != null) {
                    ImageIcon original = new ImageIcon(resource);
                    
                    // For small icons, use fast scaling
                    if (width <= 32 && height <= 32) {
                        Image scaled = original.getImage().getScaledInstance(
                            width, height, Image.SCALE_SMOOTH);
                        return new ImageIcon(scaled);
                    }
                    
                    // For larger icons, use quality scaling but don't cache
                    BufferedImage buffered = toBufferedImage(original.getImage());
                    Image scaled = scaleImageFast(buffered, width, height);
                    return new ImageIcon(scaled);
                }
            }
            return null;
        } catch (Exception e) {
            System.err.println("Error loading icon: " + filename);
            return null;
        }
    }
    
    // ============================================================================
    // FAST IMAGE SCALING - Optimized for speed, not caching
    // ============================================================================
    
    /**
     * Fast single-pass scaling - minimal memory usage
     */
    private static Image scaleImageFast(BufferedImage original, int targetWidth, int targetHeight) {
        // Direct scaling for small changes
        if (Math.abs(original.getWidth() - targetWidth) < original.getWidth() / 4) {
            return scaleSinglePass(original, targetWidth, targetHeight);
        }
        
        // Multi-pass only for large downscaling
        if (targetWidth < original.getWidth() / 3) {
            return scaleMultiPassMemoryEfficient(original, targetWidth, targetHeight);
        }
        
        return scaleSinglePass(original, targetWidth, targetHeight);
    }
    
    private static Image scaleSinglePass(BufferedImage original, int width, int height) {
        BufferedImage scaled = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = scaled.createGraphics();
        
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_SPEED);
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        g2d.drawImage(original, 0, 0, width, height, null);
        g2d.dispose();
        
        return scaled;
    }
    
    private static Image scaleMultiPassMemoryEfficient(BufferedImage original, int targetWidth, int targetHeight) {
        int currentWidth = original.getWidth();
        int currentHeight = original.getHeight();
        BufferedImage current = original;
        
        // Scale down in 50% steps, disposing old images immediately
        while (currentWidth > targetWidth * 2) {
            currentWidth /= 2;
            currentHeight /= 2;
            
            BufferedImage temp = new BufferedImage(currentWidth, currentHeight, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2d = temp.createGraphics();
            g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g2d.drawImage(current, 0, 0, currentWidth, currentHeight, null);
            g2d.dispose();
            
            // Free previous image if not the original
            if (current != original) {
                current.flush();
            }
            current = temp;
        }
        
        // Final scaling
        Image result = scaleSinglePass(current, targetWidth, targetHeight);
        if (current != original) {
            current.flush();
        }
        
        return result;
    }
    
    private static BufferedImage toBufferedImage(Image img) {
        if (img instanceof BufferedImage) {
            return (BufferedImage) img;
        }
        
        BufferedImage buffered = new BufferedImage(
            img.getWidth(null), img.getHeight(null), BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = buffered.createGraphics();
        g2d.drawImage(img, 0, 0, null);
        g2d.dispose();
        
        return buffered;
    }
    
    // ============================================================================
    // MEMORY-EFFICIENT IMAGE VIEWER
    // ============================================================================
    
    /**
     * Opens image with aggressive memory management
     * - Limits maximum dimensions
     * - Scales down immediately to fit screen
     * - Only keeps scaled version in memory
     */
    public static ImageViewResult openImageFile(File file, int containerWidth, int containerHeight) {
        try {
            // Load image with size limits
            ImageIcon originalIcon = new ImageIcon(file.getAbsolutePath());
            Image originalImage = originalIcon.getImage();
            
            int originalWidth = originalImage.getWidth(null);
            int originalHeight = originalImage.getHeight(null);
            
            // Protect against huge images
            if (originalWidth > MAX_IMAGE_DIMENSION || originalHeight > MAX_IMAGE_DIMENSION) {
                System.out.println("⚠ Large image detected: " + originalWidth + "x" + originalHeight);
                System.out.println("  Scaling down to save memory...");
                
                double scale = Math.min(
                    (double) MAX_IMAGE_DIMENSION / originalWidth,
                    (double) MAX_IMAGE_DIMENSION / originalHeight
                );
                
                originalWidth = (int) (originalWidth * scale);
                originalHeight = (int) (originalHeight * scale);
                
                // Create scaled-down version as the "original"
                BufferedImage temp = toBufferedImage(originalImage);
                originalImage = scaleImageFast(temp, originalWidth, originalHeight);
                temp.flush(); // Free the huge original
            }
            
            // Calculate display size
            int maxWidth = containerWidth - 100;
            int maxHeight = containerHeight - 150;
            
            double scale = 1.0;
            if (originalWidth > maxWidth || originalHeight > maxHeight) {
                double widthRatio = (double) maxWidth / originalWidth;
                double heightRatio = (double) maxHeight / originalHeight;
                scale = Math.min(widthRatio, heightRatio);
            }
            
            int displayWidth = (int) (originalWidth * scale);
            int displayHeight = (int) (originalHeight * scale);
            
            // Create display image
            BufferedImage buffered = toBufferedImage(originalImage);
            Image displayImage = scaleImageFast(buffered, displayWidth, displayHeight);
            
            ImageIcon displayIcon = new ImageIcon(displayImage);
            JLabel imageLabel = new JLabel(displayIcon);
            imageLabel.setHorizontalAlignment(JLabel.CENTER);
            
            JScrollPane scrollPane = new JScrollPane(imageLabel);
            scrollPane.getVerticalScrollBar().setUnitIncrement(16);
            scrollPane.getHorizontalScrollBar().setUnitIncrement(16);
            
            // Keep only the scaled version in memory, not the original
            final double[] zoomLevel = {scale};
            final int[] baseDimensions = {originalWidth, originalHeight};
            
            setupMemoryEfficientZoom(imageLabel, file, baseDimensions, zoomLevel);
            
            // Free the buffered image
            buffered.flush();
            
            return new ImageViewResult(scrollPane, imageLabel, null, zoomLevel, file, baseDimensions);
            
        } catch (Exception e) {
            throw new RuntimeException("Unable to open image: " + e.getMessage(), e);
        }
    }
    
    /**
     * Memory-efficient zoom that reloads from disk on demand
     */
    private static void setupMemoryEfficientZoom(JLabel imageLabel, File sourceFile, 
                                                int[] baseDimensions, double[] zoomLevel) {
        InputMap im = imageLabel.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW);
        ActionMap am = imageLabel.getActionMap();
        
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_PLUS, InputEvent.CTRL_DOWN_MASK), "zoomIn");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_EQUALS, InputEvent.CTRL_DOWN_MASK), "zoomIn");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_ADD, InputEvent.CTRL_DOWN_MASK), "zoomIn");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_MINUS, InputEvent.CTRL_DOWN_MASK), "zoomOut");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_SUBTRACT, InputEvent.CTRL_DOWN_MASK), "zoomOut");
        
        am.put("zoomIn", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                zoomLevel[0] *= 1.1;
                resizeImageFromFile(imageLabel, sourceFile, baseDimensions, zoomLevel[0]);
            }
        });
        
        am.put("zoomOut", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                zoomLevel[0] /= 1.1;
                resizeImageFromFile(imageLabel, sourceFile, baseDimensions, zoomLevel[0]);
            }
        });
    }
    
    /**
     * Reload and scale image on demand - doesn't keep original in memory
     */
    private static void resizeImageFromFile(JLabel label, File sourceFile, int[] baseDimensions, double scale) {
        SwingUtilities.invokeLater(() -> {
            try {
                int newWidth = (int) (baseDimensions[0] * scale);
                int newHeight = (int) (baseDimensions[1] * scale);
                
                if (newWidth <= 10 || newHeight <= 10) return;
                
                // Check if we have a soft reference cached
                String cacheKey = sourceFile.getAbsolutePath();
                SoftReference<BufferedImage> ref = IMAGE_CACHE.get(cacheKey);
                BufferedImage original = (ref != null) ? ref.get() : null;
                
                // Reload if not cached or GC'd
                if (original == null) {
                    ImageIcon icon = new ImageIcon(sourceFile.getAbsolutePath());
                    original = toBufferedImage(icon.getImage());
                    
                    // Store with SoftReference - GC can clear if memory is low
                    IMAGE_CACHE.put(cacheKey, new SoftReference<>(original));
                }
                
                // Scale to requested size
                Image scaled = scaleImageFast(original, newWidth, newHeight);
                label.setIcon(new ImageIcon(scaled));
                label.revalidate();
                label.repaint();
                
            } catch (Exception e) {
                System.err.println("Error resizing image: " + e.getMessage());
            }
        });
    }
    
    @Deprecated
    public static void resizeImage(JLabel label, Image original, double scale) {
        int newWidth = (int) (original.getWidth(null) * scale);
        int newHeight = (int) (original.getHeight(null) * scale);
        if (newWidth <= 10 || newHeight <= 10) return;
        
        BufferedImage buffered = toBufferedImage(original);
        Image scaled = scaleImageFast(buffered, newWidth, newHeight);
        label.setIcon(new ImageIcon(scaled));
        label.revalidate();
        label.repaint();
    }
    
    // ============================================================================
    // FALLBACK ICON GENERATORS - Minimal memory
    // ============================================================================
    
    private static Icon createDefaultCloseIcon(int width, int height) {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        g2.setColor(new Color(255, 100, 100));
        g2.fillOval(1, 1, width-2, height-2);
        
        g2.setColor(Color.WHITE);
        g2.setStroke(new BasicStroke(Math.max(1.5f, width / 10f)));
        int pad = width / 4;
        g2.drawLine(pad, pad, width-pad, height-pad);
        g2.drawLine(width-pad, pad, pad, height-pad);
        
        g2.dispose();
        return new ImageIcon(image);
    }
    
    private static Icon createDefaultNewTabIcon(int width, int height) {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        g2.setColor(new Color(100, 200, 100));
        g2.fillOval(1, 1, width-2, height-2);
        
        g2.setColor(Color.WHITE);
        g2.setStroke(new BasicStroke(Math.max(1.5f, width / 10f)));
        g2.drawLine(width/2, width/4, width/2, 3*width/4);
        g2.drawLine(width/4, height/2, 3*width/4, height/2);
        
        g2.dispose();
        return new ImageIcon(image);
    }
    
    private static Icon createDefaultTableIcon(int width, int height) {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        g2.setColor(new Color(100, 150, 255));
        g2.fillRoundRect(0, 0, width, height, width/8, height/8);
        
        g2.setColor(Color.WHITE);
        g2.setStroke(new BasicStroke(Math.max(1f, width / 16f)));
        
        int pad = 2;
        g2.drawLine(width/3, pad, width/3, height-pad);
        g2.drawLine(2*width/3, pad, 2*width/3, height-pad);
        g2.drawLine(pad, height/3, width-pad, height/3);
        g2.drawLine(pad, 2*height/3, width-pad, 2*height/3);
        
        g2.dispose();
        return new ImageIcon(image);
    }
    
    // ============================================================================
    // LEGACY & UTILITY METHODS
    // ============================================================================
    
    @Deprecated
    public static Icon createScaledIcon(String resourcePath, int width, int height) {
        return loadIconOptimized(resourcePath.replace("/icons/", "").replace("/images/", ""), width, height);
    }
    
    public static Icon createScaledIcon(ImageIcon originalIcon, int width, int height) {
        BufferedImage buffered = toBufferedImage(originalIcon.getImage());
        Image scaled = scaleImageFast(buffered, width, height);
        return new ImageIcon(scaled);
    }
    
    public static Icon createScaledIconAspectRatio(String resourcePath, int maxWidth, int maxHeight) {
        String filename = resourcePath.substring(resourcePath.lastIndexOf('/') + 1);
        Icon originalIcon = loadIconOptimized(filename, -1, -1);
        
        if (originalIcon == null) {
            return createPlaceholderIcon(maxWidth, maxHeight, Color.LIGHT_GRAY);
        }
        
        ImageIcon imgIcon = (ImageIcon) originalIcon;
        int originalWidth = imgIcon.getIconWidth();
        int originalHeight = imgIcon.getIconHeight();
        
        double ratio = Math.min((double) maxWidth / originalWidth, (double) maxHeight / originalHeight);
        int newWidth = (int) (originalWidth * ratio);
        int newHeight = (int) (originalHeight * ratio);
        
        BufferedImage buffered = toBufferedImage(imgIcon.getImage());
        Image scaled = scaleImageFast(buffered, newWidth, newHeight);
        return new ImageIcon(scaled);
    }
    
    public static Icon createPlaceholderIcon(int width, int height, Color color) {
        BufferedImage placeholder = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = placeholder.createGraphics();
        g2.setColor(color);
        g2.fillRect(0, 0, width, height);
        g2.setColor(Color.DARK_GRAY);
        g2.drawRect(0, 0, width - 1, height - 1);
        g2.dispose();
        return new ImageIcon(placeholder);
    }
    
    public static Dimension getImageDimensions(File file) {
        try {
            ImageIcon icon = new ImageIcon(file.getAbsolutePath());
            return new Dimension(icon.getIconWidth(), icon.getIconHeight());
        } catch (Exception e) {
            return new Dimension(0, 0);
        }
    }
    
    public static boolean isSupportedImageFile(File file) {
        if (file == null || !file.exists()) return false;
        String name = file.getName().toLowerCase();
        return name.endsWith(".jpg") || name.endsWith(".jpeg") || 
               name.endsWith(".png") || name.endsWith(".gif") || 
               name.endsWith(".bmp") || name.endsWith(".webp");
    }
    
    /**
     * Aggressively clear all caches and suggest GC
     */
    public static void clearCache() {
        ICON_CACHE.clear();
        IMAGE_CACHE.clear();
        System.gc(); // Suggest garbage collection
        System.out.println("✓ Image cache cleared");
    }
    
    /**
     * Get current cache memory usage estimate
     */
    public static String getCacheStats() {
        int iconCount = ICON_CACHE.size();
        int imageCount = IMAGE_CACHE.size();
        return String.format("Icons cached: %d | Images cached: %d", iconCount, imageCount);
    }
    
    // ============================================================================
    // FUNCTIONAL INTERFACE
    // ============================================================================
    
    @FunctionalInterface
    private interface IconGenerator {
        Icon generate(int width, int height);
    }
    
    // ============================================================================
    // ENHANCED RESULT CLASS - Stores file reference instead of full image
    // ============================================================================
    
    public static class ImageViewResult {
        private final JScrollPane scrollPane;
        private final JLabel imageLabel;
        private final Image originalImage; // Kept for backward compatibility
        private final double[] zoomLevel;
        private final File sourceFile;
        private final int[] baseDimensions;
        
        public ImageViewResult(JScrollPane scrollPane, JLabel imageLabel, Image originalImage, 
                             double[] zoomLevel, File sourceFile, int[] baseDimensions) {
            this.scrollPane = scrollPane;
            this.imageLabel = imageLabel;
            this.originalImage = originalImage;
            this.zoomLevel = zoomLevel;
            this.sourceFile = sourceFile;
            this.baseDimensions = baseDimensions;
        }
        
        // Old constructor for backward compatibility
        public ImageViewResult(JScrollPane scrollPane, JLabel imageLabel, Image originalImage, double[] zoomLevel) {
            this(scrollPane, imageLabel, originalImage, zoomLevel, null, null);
        }
        
        public JScrollPane getScrollPane() { return scrollPane; }
        public JLabel getImageLabel() { return imageLabel; }
        public Image getOriginalImage() { return originalImage; }
        public double getZoomLevel() { return zoomLevel[0]; }
        public void setZoomLevel(double level) { zoomLevel[0] = level; }
        public File getSourceFile() { return sourceFile; }
        public int[] getBaseDimensions() { return baseDimensions; }
        
        /**
         * Free memory used by this image view
         */
        public void dispose() {
            if (originalImage != null) {
                originalImage.flush();
            }
            if (imageLabel.getIcon() instanceof ImageIcon) {
                ((ImageIcon) imageLabel.getIcon()).getImage().flush();
            }
        }
    }
}