// SidebarPanel.java
package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.geom.RoundRectangle2D;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;

import org.bioparse.cleaning.Constants;

public class SidebarPanel extends JPanel {
    private static final long serialVersionUID = 1L;
    private JButton activeButton = null;
    private final Color[] GRADIENT_COLORS = {
        new Color(44, 62, 80),      // Dark blue
        new Color(52, 73, 94),      // Slightly lighter
        new Color(41, 128, 185)     // Accent blue
    };
    
    private final Color HOVER_COLOR = new Color(255, 255, 255, 25);
    private final Color ACTIVE_COLOR = new Color(52, 152, 219, 120);
    private final Color GLOW_COLOR = new Color(52, 152, 219, 60);
    private final Color TEXT_WHITE = new Color(240, 240, 245);
    private final Color SUBTEXT_COLOR = new Color(180, 190, 200);
    
    public SidebarPanel(ActionListener homeListener, ActionListener configListener,
                       ActionListener dataListener, ActionListener reportListener,
                       ActionListener vizListener) {
        setLayout(new BorderLayout());
        setPreferredSize(new Dimension(300, Constants.SIDEBAR_SIZE.height));
        
        // Main content panel
        JPanel contentPanel = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                
                // Create gradient background
                GradientPaint gradient = new GradientPaint(
                    0, 0, GRADIENT_COLORS[0], 
                    getWidth(), getHeight(), GRADIENT_COLORS[1]
                );
                g2d.setPaint(gradient);
                g2d.fillRect(0, 0, getWidth(), getHeight());
                
                // Add subtle pattern overlay
                g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, 0.03f));
                g2d.setColor(Color.WHITE);
                for (int y = 0; y < getHeight(); y += 20) {
                    for (int x = 0; x < getWidth(); x += 20) {
                        g2d.fillOval(x, y, 2, 2);
                    }
                }
                g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, 1f));
                
                // Add subtle glow at top
                GradientPaint topGlow = new GradientPaint(
                    0, 0, new Color(52, 152, 219, 30),
                    0, 100, new Color(52, 152, 219, 0)
                );
                g2d.setPaint(topGlow);
                g2d.fillRect(0, 0, getWidth(), 100);
            }
        };
        
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
        contentPanel.setBorder(new EmptyBorder(25, 20, 25, 20));
        
        setupHeader(contentPanel);
        setupButtons(contentPanel, homeListener, configListener, dataListener, reportListener, vizListener);
        setupFooter(contentPanel);
        
        JScrollPane scrollPane = new JScrollPane(contentPanel);
        scrollPane.setBorder(null);
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scrollPane.getVerticalScrollBar().setUI(new ModernScrollBarUI());
        scrollPane.setViewportBorder(null);
        scrollPane.getViewport().setOpaque(false);
        scrollPane.setOpaque(false);
        
        add(scrollPane, BorderLayout.CENTER);
    }
    
    private void setupHeader(JPanel parent) {
        JPanel headerPanel = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                
                // Add decorative circle behind logo
                g2d.setColor(new Color(52, 152, 219, 20));
                g2d.fillOval(getWidth() - 80, -30, 100, 100);
            }
        };
        
        headerPanel.setLayout(new BoxLayout(headerPanel, BoxLayout.Y_AXIS));
        headerPanel.setOpaque(false);
        headerPanel.setBorder(new EmptyBorder(0, 10, 40, 10));
        headerPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        // App logo/icon with gradient
        JLabel logoIcon = new JLabel("📊") {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                
                // Draw background circle
                g2d.setColor(new Color(52, 152, 219, 40));
                g2d.fillOval(0, 0, getWidth(), getHeight());
                
                // Draw gradient border
                GradientPaint borderGradient = new GradientPaint(
                    0, 0, new Color(52, 152, 219, 200),
                    getWidth(), getHeight(), new Color(41, 128, 185, 200)
                );
                g2d.setPaint(borderGradient);
                g2d.setStroke(new BasicStroke(2));
                g2d.drawOval(1, 1, getWidth()-2, getHeight()-2);
                
                super.paintComponent(g);
            }
        };
        
        logoIcon.setFont(new Font("Segoe UI Emoji", Font.PLAIN, 28));
        logoIcon.setHorizontalAlignment(SwingConstants.CENTER);
        logoIcon.setPreferredSize(new Dimension(60, 60));
        logoIcon.setMaximumSize(new Dimension(60, 60));
        logoIcon.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        // App name with gradient text
        JLabel appName = new JLabel("<html><div style='text-align: left;'><b style='font-size: 24px;'>Attendance</b><br>" +
                                   "<span style='font-size: 24px; color: #3498db;'>Pro</span></div></html>");
        appName.setFont(new Font("Segoe UI", Font.BOLD, 24));
        appName.setForeground(TEXT_WHITE);
        appName.setAlignmentX(Component.LEFT_ALIGNMENT);
        appName.setBorder(new EmptyBorder(15, 0, 5, 0));
        
        // Subtitle
        JLabel subtitle = new JLabel("Data Intelligence Platform");
        subtitle.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        subtitle.setForeground(SUBTEXT_COLOR);
        subtitle.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        headerPanel.add(logoIcon);
        headerPanel.add(Box.createRigidArea(new Dimension(0, 15)));
        headerPanel.add(appName);
        headerPanel.add(subtitle);
        
        parent.add(headerPanel);
    }
    
    private void setupButtons(JPanel parent, ActionListener homeListener, 
                             ActionListener configListener, ActionListener dataListener,
                             ActionListener reportListener, ActionListener vizListener) {
        JPanel buttonPanel = new JPanel();
        buttonPanel.setLayout(new BoxLayout(buttonPanel, BoxLayout.Y_AXIS));
        buttonPanel.setOpaque(false);
        buttonPanel.setBorder(new EmptyBorder(0, 0, 30, 0));
        buttonPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        // Section Label
        JLabel sectionLabel = new JLabel("NAVIGATION");
        sectionLabel.setFont(new Font("Segoe UI", Font.BOLD, 10));
        sectionLabel.setForeground(new Color(150, 160, 180));
        sectionLabel.setBorder(new EmptyBorder(0, 15, 15, 0));
        sectionLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        buttonPanel.add(sectionLabel);
        
        addModernSidebarButton("Dashboard", "🏠", homeListener, buttonPanel);
        addModernSidebarButton("Configuration", "⚙️", configListener, buttonPanel);
        addModernSidebarButton("Data Viewer", "📊", dataListener, buttonPanel);
        addModernSidebarButton("Reports", "📑", reportListener, buttonPanel);
        addModernSidebarButton("Analytics", "📈", vizListener, buttonPanel);
        
        parent.add(buttonPanel);
    }
    
    private void addModernSidebarButton(String text, String icon, ActionListener listener, JPanel parent) {
        JButton button = new JButton() {
            private boolean hover = false;
            private boolean active = false;
            
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                
                int width = getWidth();
                int height = getHeight();
                
                // Draw background based on state
                if (active) {
                    // Active state with gradient and glow
                    GradientPaint activeGradient = new GradientPaint(
                        0, 0, new Color(52, 152, 219, 150),
                        0, height, new Color(41, 128, 185, 150)
                    );
                    g2d.setPaint(activeGradient);
                    g2d.fillRoundRect(0, 0, width, height, 15, 15);
                    
                    // Glow effect
                    g2d.setColor(GLOW_COLOR);
                    g2d.setStroke(new BasicStroke(2));
                    g2d.drawRoundRect(1, 1, width-2, height-2, 15, 15);
                    
                } else if (hover) {
                    // Hover state
                    g2d.setColor(HOVER_COLOR);
                    g2d.fillRoundRect(0, 0, width, height, 15, 15);
                }
                
                // Draw icon with animation effect
                int iconSize = 20;
                int iconY = (height - iconSize) / 2;
                
                if (hover || active) {
                    // Pulsing effect for icon
                    g2d.setColor(new Color(255, 255, 255, 220));
                    g2d.fillOval(20, iconY-5, iconSize+10, iconSize+10);
                }
                
                // Draw icon
                g2d.setFont(new Font("Segoe UI Emoji", Font.PLAIN, 18));
                g2d.setColor(active ? Color.WHITE : (hover ? Color.WHITE : SUBTEXT_COLOR));
                FontMetrics fm = g2d.getFontMetrics();
                int iconX = 25;
                g2d.drawString(icon, iconX, iconY + fm.getAscent() - 5);
                
                // Draw text
                g2d.setFont(new Font("Segoe UI", Font.PLAIN, 14));
                g2d.setColor(active ? Color.WHITE : TEXT_WHITE);
                g2d.drawString(text, 60, height / 2 + 5);
                
                // Draw active indicator arrow
                if (active) {
                    g2d.setColor(Color.WHITE);
                    int[] xPoints = {width - 15, width - 5, width - 15};
                    int[] yPoints = {height/2 - 5, height/2, height/2 + 5};
                    g2d.fillPolygon(xPoints, yPoints, 3);
                }
            }
        };
        
        button.setContentAreaFilled(false);
        button.setBorderPainted(false);
        button.setFocusPainted(false);
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setToolTipText("Open " + text);
        
        Dimension buttonSize = new Dimension(260, 50);
        button.setMaximumSize(buttonSize);
        button.setPreferredSize(buttonSize);
        button.setMinimumSize(buttonSize);
        button.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        button.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                button.putClientProperty("hover", true);
                button.repaint();
            }
            
            @Override
            public void mouseExited(MouseEvent e) {
                button.putClientProperty("hover", false);
                button.repaint();
            }
        });
        
        button.addActionListener(e -> {
            if (activeButton != null) {
                activeButton.putClientProperty("active", false);
                activeButton.repaint();
            }
            button.putClientProperty("active", true);
            activeButton = button;
            button.repaint();
            listener.actionPerformed(e);
        });
        
        parent.add(button);
        parent.add(Box.createRigidArea(new Dimension(0, 8)));
    }
    
    private void setupFooter(JPanel parent) {
        JPanel footerPanel = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                
                // Draw decorative line
                g2d.setColor(new Color(255, 255, 255, 20));
                g2d.fillRoundRect(20, 0, getWidth()-40, 1, 5, 5);
            }
        };
        
        footerPanel.setLayout(new BoxLayout(footerPanel, BoxLayout.Y_AXIS));
        footerPanel.setOpaque(false);
        footerPanel.setBorder(new EmptyBorder(20, 20, 0, 20));
        footerPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        
        // User profile
        JPanel userPanel = new JPanel();
        userPanel.setLayout(new BorderLayout(10, 0));
        userPanel.setOpaque(false);
        userPanel.setMaximumSize(new Dimension(260, 60));
        
        // Avatar
        JLabel avatar = new JLabel("👤") {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                
                // Draw circular avatar background
                g2d.setColor(new Color(255, 255, 255, 15));
                g2d.fillOval(0, 0, getWidth(), getHeight());
                
                super.paintComponent(g);
            }
        };
        
        avatar.setFont(new Font("Segoe UI Emoji", Font.PLAIN, 20));
        avatar.setHorizontalAlignment(SwingConstants.CENTER);
        avatar.setPreferredSize(new Dimension(40, 40));
        
        // User info
        JPanel infoPanel = new JPanel();
        infoPanel.setLayout(new BoxLayout(infoPanel, BoxLayout.Y_AXIS));
        infoPanel.setOpaque(false);
        
        JLabel userName = new JLabel("Admin User");
        userName.setFont(new Font("Segoe UI", Font.BOLD, 13));
        userName.setForeground(TEXT_WHITE);
        
        JLabel userRole = new JLabel("Administrator");
        userRole.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        userRole.setForeground(SUBTEXT_COLOR);
        
        infoPanel.add(userName);
        infoPanel.add(userRole);
        
        userPanel.add(avatar, BorderLayout.WEST);
        userPanel.add(infoPanel, BorderLayout.CENTER);
        
        // Version
        JLabel versionLabel = new JLabel("v2.5.1");
        versionLabel.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        versionLabel.setForeground(new Color(150, 160, 180, 150));
        versionLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        versionLabel.setBorder(new EmptyBorder(10, 20, 0, 0));
        
        footerPanel.add(userPanel);
        footerPanel.add(versionLabel);
        
        parent.add(Box.createVerticalGlue());
        parent.add(footerPanel);
    }
    
    // Custom ScrollBar UI
    class ModernScrollBarUI extends javax.swing.plaf.basic.BasicScrollBarUI {
        private final Dimension THUMB_SIZE = new Dimension(4, 0);
        private final Color TRACK_COLOR = new Color(255, 255, 255, 10);
        private final Color THUMB_COLOR = new Color(255, 255, 255, 40);
        private final Color THUMB_HOVER_COLOR = new Color(255, 255, 255, 60);
        
        @Override
        protected void paintThumb(Graphics g, JComponent c, Rectangle thumbBounds) {
            if (thumbBounds.isEmpty() || !scrollbar.isEnabled()) {
                return;
            }
            
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            
            Color color = isThumbRollover() ? THUMB_HOVER_COLOR : THUMB_COLOR;
            g2.setColor(color);
            g2.fillRoundRect(thumbBounds.x + 2, thumbBounds.y + 4, 
                           thumbBounds.width - 4, thumbBounds.height - 8, 4, 4);
            g2.dispose();
        }
        
        @Override
        protected void paintTrack(Graphics g, JComponent c, Rectangle trackBounds) {
            g.setColor(TRACK_COLOR);
            g.fillRect(trackBounds.x, trackBounds.y, trackBounds.width, trackBounds.height);
        }
        
        @Override
        protected JButton createDecreaseButton(int orientation) {
            return createZeroButton();
        }
        
        @Override
        protected JButton createIncreaseButton(int orientation) {
            return createZeroButton();
        }
        
        private JButton createZeroButton() {
            JButton button = new JButton();
            button.setPreferredSize(new Dimension(0, 0));
            button.setMinimumSize(new Dimension(0, 0));
            button.setMaximumSize(new Dimension(0, 0));
            return button;
        }
        
        @Override
        protected Dimension getMinimumThumbSize() {
            return THUMB_SIZE;
        }
    }
}