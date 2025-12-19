// UIComponentFactory.java
package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import org.bioparse.cleaning.Constants;

public class UIComponentFactory {
    
    public static JPanel createCardPanel() {
        JPanel card = new JPanel();
        card.setBackground(Color.WHITE);
        card.setBorder(new CompoundBorder(
            new LineBorder(Constants.BORDER_COLOR, 1, true),
            new EmptyBorder(15, 15, 15, 15)
        ));
        return card;
    }
    
    public static JLabel createStyledLabel(String text) {
        JLabel label = new JLabel(text);
        label.setFont(Constants.FONT_BODY);
        label.setForeground(Constants.TEXT_PRIMARY);
        return label;
    }
    
    public static JTextField createStyledTextField(int columns) {
        JTextField field = new JTextField(columns);
        field.setFont(Constants.FONT_BODY);
        field.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(Constants.BORDER_COLOR, 1),
            new EmptyBorder(8, 10, 8, 10)
        ));
        return field;
    }
    
    public static JComboBox<String> createStyledComboBox() {
        JComboBox<String> combo = new JComboBox<>();
        combo.setFont(Constants.FONT_BODY);
        combo.setBackground(Color.WHITE);
        combo.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(Constants.BORDER_COLOR, 1),
            new EmptyBorder(5, 10, 5, 10)
        ));
        return combo;
    }
    
    public static JButton createPrimaryButton(String text) {
        return createButton(text, Constants.PRIMARY_COLOR, Constants.SECONDARY_COLOR, Color.WHITE);
    }
    
    public static JButton createSecondaryButton(String text) {
        return createButton(text, Color.WHITE, new Color(245, 245, 245), Constants.TEXT_PRIMARY);
    }
    
    public static JButton createSuccessButton(String text) {
        return createButton(text, Constants.SUCCESS_COLOR, new Color(46, 204, 113), Color.WHITE);
    }
    
    private static JButton createButton(String text, Color bgColor, Color hoverBg, Color fgColor) {
        JButton button = new JButton(text);
        button.setFont(Constants.FONT_BUTTON);
        button.setBackground(bgColor);
        button.setForeground(fgColor);
        button.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(bgColor, 1, true),
            new EmptyBorder(10, 15, 10, 15)
        ));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setFocusPainted(false);
        
        button.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent evt) {
                button.setBackground(hoverBg);
            }
            
            public void mouseExited(MouseEvent evt) {
                button.setBackground(bgColor);
            }
        });
        
        return button;
    }
}