package com.example.notepadpp.media;
import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Image;
import java.awt.Window;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.Hashtable;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JSlider;
import javax.swing.JWindow;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.event.ChangeEvent;
import uk.co.caprica.vlcj.factory.discovery.NativeDiscovery;
import uk.co.caprica.vlcj.player.base.MediaPlayer;
import uk.co.caprica.vlcj.player.base.MediaPlayerEventAdapter;
import uk.co.caprica.vlcj.player.component.EmbeddedMediaPlayerComponent;
public class VLCJVideoPlayerPanel extends JPanel {private final EmbeddedMediaPlayerComponent mediaPlayerComponent;private boolean isSeeking = false;private boolean isFullScreen = false;public VLCJVideoPlayerPanel(File file) {
        setLayout(new BorderLayout());
        System.setProperty("jna.library.path", "lib/vlc");
        System.setProperty("VLC_PLUGIN_PATH", "lib/vlc/plugins");
        new NativeDiscovery().discover();
        mediaPlayerComponent = new EmbeddedMediaPlayerComponent();
        add(mediaPlayerComponent, BorderLayout.CENTER);
        JPanel basicControls = new JPanel(new BorderLayout());
        basicControls.setBackground(UIManager.getColor("Panel.background"));
        JPanel leftButtons = new JPanel();
        leftButtons.setBackground(UIManager.getColor("Panel.background"));
        JButton playPause = new JButton("⏯");
        JButton stop = new JButton("⏹");
        leftButtons.add(playPause);
        leftButtons.add(stop);
        leftButtons.add(new JLabel("  🔊 "));
        JButton fullScreenButton = new JButton("Full Screen");
        fullScreenButton.addActionListener((ActionEvent e) -> {
            toggleFullScreen();
        });
        JSlider volume = new JSlider(0, 100, 80);
        volume.setPreferredSize(new Dimension(120, volume.getPreferredSize().height));
        volume.setPaintTicks(true);
        volume.setMajorTickSpacing(20);
        volume.setMinorTickSpacing(5);
        volume.setPaintLabels(true);
        leftButtons.add(volume);
        basicControls.add(leftButtons, BorderLayout.WEST);
        JSlider seekBar = new JSlider(0, 100, 0);
        basicControls.add(seekBar, BorderLayout.CENTER);
        JPanel extraControls = new JPanel();
        extraControls.setBackground(UIManager.getColor("Panel.background"));
        JButton rewind = new JButton("⏪");
        JButton fastForward = new JButton("⏩");
        JButton muteBtn = new JButton("Mute");
        extraControls.add(rewind);
        extraControls.add(fastForward);
        extraControls.add(muteBtn);
        extraControls.add(new JLabel(" Speed: "));
        JSlider speedSlider = new JSlider(50, 150, 100);
        speedSlider.setPreferredSize(new Dimension(120, speedSlider.getPreferredSize().height));
        speedSlider.setPaintTicks(true);
        speedSlider.setMajorTickSpacing(25);
        speedSlider.setMinorTickSpacing(5);
        Hashtable<Integer, JLabel> speedLabels = new Hashtable<>();
        speedLabels.put(50, new JLabel("50%"));
        speedLabels.put(75, new JLabel("75%"));
        speedLabels.put(100, new JLabel("100%"));
        speedLabels.put(125, new JLabel("125%"));
        speedLabels.put(150, new JLabel("150%"));
        speedLabels.put(200, new JLabel("200%"));
        speedSlider.setLabelTable(speedLabels);
        speedSlider.setPaintLabels(true);
        extraControls.add(speedSlider);
        final JLabel timeLabel = new JLabel("00:00 / 00:00");
        extraControls.add(timeLabel);
        JPanel controlContainer = new JPanel();
        controlContainer.setLayout(new BoxLayout(controlContainer, BoxLayout.Y_AXIS));
        controlContainer.setBackground(UIManager.getColor("Panel.background"));
        controlContainer.add(basicControls);
        controlContainer.add(extraControls);
        controlContainer.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        add(controlContainer, BorderLayout.SOUTH);
        MediaPlayer player = mediaPlayerComponent.mediaPlayer();
        player.audio().setVolume(80);
        SwingUtilities.invokeLater(() -> player.media().play(file.getAbsolutePath()));
        playPause.addActionListener((ActionEvent e) -> {
            if (player.status().isPlaying()) {
                player.controls().pause();
            } else {
                player.controls().play();
            }
        });
        stop.addActionListener((ActionEvent e) -> player.controls().stop());
        volume.addChangeListener((ChangeEvent e) -> player.audio().setVolume(volume.getValue()));
        seekBar.addChangeListener((ChangeEvent e) -> {
            if (seekBar.getValueIsAdjusting()) {
                isSeeking = true;
            } else {
                if (isSeeking) {
                    long duration = player.media().info().duration();
                    int value = seekBar.getValue();
                    long time = (duration * value) / 100;
                    player.controls().setTime(time);
                    isSeeking = false;
                }
            }
        });
        rewind.addActionListener((ActionEvent e) -> {
            long currentTime = player.status().time();
            player.controls().setTime(Math.max(0, currentTime - 10000));
        });
        fastForward.addActionListener((ActionEvent e) -> {
            long currentTime = player.status().time();
            long duration = player.media().info().duration();
            player.controls().setTime(Math.min(duration, currentTime + 10000));
        });
        muteBtn.addActionListener((ActionEvent e) -> {
            boolean mute = player.audio().isMute();
            player.audio().setMute(!mute);
            muteBtn.setText(mute ? "Mute" : "Unmute");
        });
        speedSlider.addChangeListener((ChangeEvent e) -> {
            float newRate = speedSlider.getValue() / 100f;
            player.controls().setRate(newRate);
        });
        player.events().addMediaPlayerEventListener(new MediaPlayerEventAdapter() {
            @Override
            public void timeChanged(MediaPlayer mediaPlayer, long newTime) {
                if (!isSeeking) {
                    long total = player.media().info().duration();
                    if (total > 0) {
                        int value = (int) ((newTime * 100) / total);
                        SwingUtilities.invokeLater(() -> {
                            seekBar.setValue(value);
                            timeLabel.setText(formatTime(newTime) + " / " + formatTime(total));
                        });
                    }
                }
            }
        });
    }
    @Override
    public void addNotify() {
        super.addNotify();
        JComponent rootPane = this.getRootPane();
        if (rootPane != null) {
            rootPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
                    .put(KeyStroke.getKeyStroke("SPACE"), "togglePause");
            rootPane.getActionMap().put("togglePause", new AbstractAction() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    MediaPlayer player = mediaPlayerComponent.mediaPlayer();
                    if (player.status().isPlaying()) {
                        player.controls().pause();
                    } else {
                        player.controls().play();
                    }
                }
            });
            rootPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
                    .put(KeyStroke.getKeyStroke("RIGHT"), "forward10");
            rootPane.getActionMap().put("forward10", new AbstractAction() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    MediaPlayer player = mediaPlayerComponent.mediaPlayer();
                    long currentTime = player.status().time();
                    long duration = player.media().info().duration();
                    long newTime = Math.min(duration, currentTime + 10000);
                    player.controls().setTime(newTime);
                }
            });
            rootPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
                    .put(KeyStroke.getKeyStroke("LEFT"), "backward10");
            rootPane.getActionMap().put("backward10", new AbstractAction() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    MediaPlayer player = mediaPlayerComponent.mediaPlayer();
                    long currentTime = player.status().time();
                    long newTime = Math.max(0, currentTime - 10000);
                    player.controls().setTime(newTime);
                }
            });
        }
    }
    public void setFullScreen(boolean fullScreen) {
        Window window = SwingUtilities.getWindowAncestor(this);
        if (window == null) {
            System.err.println("setFullScreen(): Panel is not yet in a window!");
            return;
        }
        if (!(window instanceof JFrame)) {
            System.err.println("setFullScreen(): The containing window is not a JFrame.");
            return;
        }
        JFrame frame = (JFrame) window;
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        if (fullScreen) {
            System.out.println("Entering full-screen mode.");
            frame.dispose();
            frame.setUndecorated(true);
            frame.setResizable(false);
            frame.setExtendedState(JFrame.MAXIMIZED_BOTH);
            frame.setVisible(true);
            frame.validate();
            try {
                Thread.sleep(100);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
            SwingUtilities.invokeLater(() -> {
                mediaPlayerComponent.repaint();
                mediaPlayerComponent.revalidate();
                mediaPlayerComponent.requestFocusInWindow();
            });
            isFullScreen = true;
        } else {
            System.out.println("Exiting full-screen mode.");
            frame.dispose();
            frame.setUndecorated(false);
            frame.setResizable(true);
            frame.setExtendedState(JFrame.NORMAL);
            frame.setVisible(true);
            frame.validate();
            try {
                Thread.sleep(100);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
            SwingUtilities.invokeLater(() -> {
                mediaPlayerComponent.repaint();
                mediaPlayerComponent.revalidate();
                mediaPlayerComponent.requestFocusInWindow();
            });
            isFullScreen = false;
        }
    }
    public void toggleFullScreen() {
        Window window = SwingUtilities.getWindowAncestor(this);
        if (window == null) {
            System.err.println("toggleFullScreen(): Panel is not yet in a window!");
            return;
        }
        if (!(window instanceof JFrame)) {
            System.err.println("toggleFullScreen(): The containing window is not a JFrame.");
            return;
        }
        JFrame frame = (JFrame) window;
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
    
        if (!isFullScreen) {
            System.out.println("Entering full-screen mode.");
            frame.dispose();
            frame.setUndecorated(true);
            frame.setResizable(false);
            frame.setExtendedState(JFrame.MAXIMIZED_BOTH);
            frame.setVisible(true);
            frame.validate();
            try {
                Thread.sleep(100);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
            SwingUtilities.invokeLater(() -> {
                mediaPlayerComponent.repaint();
                mediaPlayerComponent.revalidate();
                mediaPlayerComponent.requestFocusInWindow();
            });
            isFullScreen = true;
        } else {
            System.out.println("Exiting full-screen mode.");
            frame.dispose();
            frame.setUndecorated(false);
            frame.setResizable(true);
            frame.setExtendedState(JFrame.NORMAL);
            frame.setVisible(true);
            frame.validate();
            try {
                Thread.sleep(100); 
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
            SwingUtilities.invokeLater(() -> {
                mediaPlayerComponent.repaint();
                mediaPlayerComponent.revalidate();
                mediaPlayerComponent.requestFocusInWindow();
            });
            isFullScreen = false;
        }
    }
    public class FullScreenHelper {
        private GraphicsDevice gd;
        private boolean isFullScreen = false;
        private JFrame frame;
        public FullScreenHelper(JFrame frame) {
            this.frame = frame;
            gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        }
        public void toggleFullScreen() {
            if (!isFullScreen) {
                frame.dispose();
                frame.setUndecorated(true);
                gd.setFullScreenWindow(frame); 
                frame.validate();
                isFullScreen = true;
            } else {
                gd.setFullScreenWindow(null);
                frame.dispose();
                frame.setUndecorated(false);
                frame.setExtendedState(JFrame.NORMAL);
                frame.setVisible(true);
                isFullScreen = false;
            }
        }
    }
    public void disposePlayer() {
        MediaPlayer player = mediaPlayerComponent.mediaPlayer();
        if (player != null && player.status().isPlaying()) {
            System.out.println("Stopping video playback...");
            player.controls().stop();
            try {
                Thread.sleep(500);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
        }
        System.out.println("Releasing video player...");
        mediaPlayerComponent.release();
    }
    private String formatTime(long millis) {
        int totalSeconds = (int) (millis / 1000);
        int minutes = totalSeconds / 60;
        int seconds = totalSeconds % 60;
        return String.format("%02d:%02d", minutes, seconds);
    }
    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(VLCJVideoPlayerPanel.class.getResource("/icons/video.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
}