package com.example.notepadpp.media;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.Hashtable;
import java.awt.Image;
import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JSlider;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.event.ChangeEvent;
import uk.co.caprica.vlcj.factory.discovery.NativeDiscovery;
import uk.co.caprica.vlcj.player.base.MediaPlayer;
import uk.co.caprica.vlcj.player.base.MediaPlayerEventAdapter;
import uk.co.caprica.vlcj.player.component.AudioPlayerComponent;

public class AudioPlayerPanel extends JPanel {

    // Instance field for the VLCJ audio player component.
    private final AudioPlayerComponent audioPlayer;

    /**
     * Constructs an AudioPlayerPanel that plays the specified audio file.
     *
     * @param file the audio file to be played
     */
    public AudioPlayerPanel(File file) {
        // Use BorderLayout for the panel.
        setLayout(new BorderLayout());
        // Set the panel background so it respects your Look & Feel.
        setBackground(UIManager.getColor("Panel.background"));

        // Top label to show the file name.
        JLabel topLabel = new JLabel("Now Playing: " + file.getName());
        add(topLabel, BorderLayout.NORTH);

        // Create a container for the controls using vertical BoxLayout.
        JPanel mainHolder = new JPanel();
        mainHolder.setLayout(new BoxLayout(mainHolder, BoxLayout.Y_AXIS));
        mainHolder.setBackground(UIManager.getColor("Panel.background"));

        // --- Basic Controls Panel ---
        JPanel controlsPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        controlsPanel.setBackground(UIManager.getColor("Panel.background"));
        JButton playBtn = new JButton("▶ Play");
        JButton pauseBtn = new JButton("⏸ Pause");
        JButton stopBtn = new JButton("⏹ Stop");
        JButton muteBtn = new JButton("Mute");

        // Volume slider with tick marks and labels.
        JSlider volumeSlider = new JSlider(0, 100, 80);
        volumeSlider.setPreferredSize(new Dimension(100, 20));
        volumeSlider.setPaintTicks(true);
        volumeSlider.setMajorTickSpacing(20);
        volumeSlider.setMinorTickSpacing(5);
        volumeSlider.setPaintLabels(true);

        controlsPanel.add(playBtn);
        controlsPanel.add(pauseBtn);
        controlsPanel.add(stopBtn);
        controlsPanel.add(muteBtn);
        controlsPanel.add(new JLabel("🔊"));
        controlsPanel.add(volumeSlider);

        // --- Progress Panel ---
        JPanel progressPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        progressPanel.setBackground(UIManager.getColor("Panel.background"));
        JSlider seekBar = new JSlider(0, 100, 0);
        seekBar.setPreferredSize(new Dimension(300, 20));
        seekBar.setPaintTicks(true);
        seekBar.setMajorTickSpacing(20);
        seekBar.setMinorTickSpacing(5);
        seekBar.setPaintLabels(true);
        JLabel timeLabel = new JLabel("00:00 / 00:00");
        progressPanel.add(seekBar);
        progressPanel.add(timeLabel);

        // --- Extra Controls Panel ---
        JPanel extraPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        extraPanel.setBackground(UIManager.getColor("Panel.background"));
        JButton rewindBtn = new JButton("⏪");
        JButton fastForwardBtn = new JButton("⏩");
        extraPanel.add(rewindBtn);
        extraPanel.add(fastForwardBtn);
        extraPanel.add(new JLabel(" Speed: "));
        // Speed slider ranging from 50% to 150%, default is 100%.
        JSlider speedSlider = new JSlider(50, 200, 100);
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
        extraPanel.add(speedSlider);
        JLabel speedValueLabel = new JLabel("100%");
        extraPanel.add(speedValueLabel);

        // Combine the sub-panels within mainHolder.
        mainHolder.add(controlsPanel);
        mainHolder.add(progressPanel);
        mainHolder.add(extraPanel);
        mainHolder.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        add(mainHolder, BorderLayout.CENTER);

        // --------------------------------------------------
        // VLCJ Audio Player Setup:
        // Set VLC native library and plugin paths (adjust these paths as needed).
        System.setProperty("jna.library.path", "lib/vlc");
        System.setProperty("VLC_PLUGIN_PATH", "lib/vlc/plugins");
        new NativeDiscovery().discover();

        // Initialize the audio player component.
        audioPlayer = new AudioPlayerComponent();
        MediaPlayer player = audioPlayer.mediaPlayer();
        player.audio().setVolume(80);

        // Start playing the audio file.
        SwingUtilities.invokeLater(() -> player.media().play(file.getAbsolutePath()));

        // --- Listeners for Controls ---
        // Play button: start or resume playback.
        playBtn.addActionListener((ActionEvent e) -> {
            if (!player.status().isPlaying()) {
                player.media().play(file.getAbsolutePath());
            } else {
                player.controls().play();
            }
        });
        // Pause button: pause playback.
        pauseBtn.addActionListener((ActionEvent e) -> player.controls().pause());
        // Stop button: stop playback.
        stopBtn.addActionListener((ActionEvent e) -> player.controls().stop());
        // Mute toggle: toggle mute status and update the button text.
        muteBtn.addActionListener((ActionEvent e) -> {
            boolean mute = player.audio().isMute();
            player.audio().setMute(!mute);
            muteBtn.setText(mute ? "Mute" : "Unmute");
        });
        // Volume slider: adjust the volume.
        volumeSlider.addChangeListener((ChangeEvent e) -> {
            int vol = volumeSlider.getValue();
            player.audio().setVolume(vol);
        });

        // Seek bar: allow user to seek using a mutable flag.
        final boolean[] isSeeking = { false };
        seekBar.addChangeListener((ChangeEvent e) -> {
            if (seekBar.getValueIsAdjusting()) {
                isSeeking[0] = true;
            } else {
                if (isSeeking[0]) {
                    long duration = player.media().info().duration();
                    int value = seekBar.getValue();
                    long time = (duration * value) / 100;
                    player.controls().setTime(time);
                    isSeeking[0] = false;
                }
            }
        });
        // Listener to update the seek bar and time label as audio plays.
        player.events().addMediaPlayerEventListener(new MediaPlayerEventAdapter() {
            @Override
            public void timeChanged(MediaPlayer mediaPlayer, long newTime) {
                if (!isSeeking[0]) {
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
        // Extra controls: Rewind (jump back 10 seconds) and Fast-forward (jump ahead 10
        // seconds).
        rewindBtn.addActionListener((ActionEvent e) -> {
            long currentTime = player.status().time();
            player.controls().setTime(Math.max(0, currentTime - 10000));
        });
        fastForwardBtn.addActionListener((ActionEvent e) -> {
            long currentTime = player.status().time();
            long duration = player.media().info().duration();
            player.controls().setTime(Math.min(duration, currentTime + 10000));
        });
        // Playback speed slider: adjust playback rate.
        speedSlider.addChangeListener((ChangeEvent e) -> {
            float newRate = speedSlider.getValue() / 100f; // Use float division.
            player.controls().setRate(newRate);
            speedValueLabel.setText(speedSlider.getValue() + "%");
        });
    }

    /**
     * Helper method to format milliseconds as a "mm:ss" time string.
     *
     * @param millis the time in milliseconds
     * @return formatted time string
     */
    private static String formatTime(long millis) {
        int totalSeconds = (int) (millis / 1000);
        int minutes = totalSeconds / 60;
        int seconds = totalSeconds % 60;
        return String.format("%02d:%02d", minutes, seconds);
    }

    /**
     * Public cleanup method to dispose the audio player.
     * Call this method when closing the tab containing the audio player.
     */
    public void disposeAudio() {
        MediaPlayer player = audioPlayer.mediaPlayer();
        if (player != null && player.status().isPlaying()) {
            System.out.println("Stopping audio playback...");
            player.controls().stop();
            try {
                Thread.sleep(500);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
        }
        System.out.println("Releasing audio player...");
        audioPlayer.release();
    }

    /**
     * Static helper method to create a new AudioPlayerPanel.
     *
     * @param file the audio file to play
     * @return a new AudioPlayerPanel instance as a JPanel
     */
    public static JPanel createAudioPanel(File file) {
        return new AudioPlayerPanel(file);
    }

    /**
     * Test the audio panel in a standalone JFrame.
     * Change the file path below to point to a valid audio file on your system.
     */
   
    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(AudioPlayerPanel.class.getResource("/icons/music.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
}