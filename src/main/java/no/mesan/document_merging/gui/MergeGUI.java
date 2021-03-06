package no.mesan.document_merging.gui;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class MergeGUI extends JFrame {

    private static final long serialVersionUID = -6058168255978115883L;

    private FileBrowserPanel filePanel;

    public MergeGUI(String defaultLayoutTemplate, String defaultContentTemplate) {
        super("Flette dokumenter");

        filePanel = new FileBrowserPanel(defaultLayoutTemplate, defaultContentTemplate);
        add(filePanel, BorderLayout.NORTH);

        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(700, 400);
        setVisible(true);
        setLocationRelativeTo(null);
    }

    private class TemplateFieldsChooserPanel extends JPanel {

    }

    private class FileBrowserPanel extends JPanel implements ActionListener {

        private JLabel lblLayoutTemplate, lblContentTemplate;
        private JFileChooser fileChooser;
        private JButton btnBrowseLayout, btnBrowseContent;
        private JTextField txtLayoutTemplate, txtContentTemplate;

        public FileBrowserPanel(String defaultLayoutTemplate, String defaultContentTemplate) {
            //TODO: Endre til gridbag layout
            setLayout(new GridLayout(2, 3));

            lblLayoutTemplate = new JLabel("Utseendemal: ");
            lblContentTemplate = new JLabel("Innholdsmal: ");

            fileChooser = new JFileChooser();
            //fileChooser.setFileFilter(new FileNameExtensionFilter("OpenOffice", "odt"));

            txtContentTemplate = defaultContentTemplate == null ?
                                 new JTextField() :
                                 new JTextField(defaultContentTemplate);
            txtLayoutTemplate = defaultLayoutTemplate == null ?
                                new JTextField() :
                                new JTextField(defaultLayoutTemplate);

            String text = "Velg fil";
            btnBrowseLayout = new JButton(text);
            btnBrowseLayout.addActionListener(this);

            btnBrowseContent = new JButton(text);
            btnBrowseContent.addActionListener(this);

            add(lblContentTemplate);
            add(txtContentTemplate);
            add(btnBrowseContent);

            add(lblLayoutTemplate);
            add(txtLayoutTemplate);
            add(btnBrowseLayout);
        }

        public void actionPerformed(ActionEvent e) {
            int choice = fileChooser.showOpenDialog(MergeGUI.this);

            if (choice == JFileChooser.APPROVE_OPTION) {

                String filePath = fileChooser.getSelectedFile().getAbsolutePath();
                if (e.getSource() == btnBrowseContent) {
                    txtContentTemplate.setText(filePath);
                } else {
                    txtLayoutTemplate.setText(filePath);
                }

            }
        }
    }

    public static void main(String[] args) {
        String path = "/media/sf_DATA_DRIVE/Documents/Jobb/Forskningsraadet/";
        String layout = path + "template.docx";
        String contents = path + "IMMedInnhold.docx";
        new MergeGUI(layout, contents);
    }
}
