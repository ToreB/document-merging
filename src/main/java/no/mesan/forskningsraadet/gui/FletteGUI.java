package no.mesan.forskningsraadet.gui;

import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

public class FletteGUI extends JFrame{
	private static final long serialVersionUID = -6058168255978115883L;
	
	private FileBrowserPanel filePanel;
	
	public FletteGUI() {
		super("Flette dokumenter");
		
		filePanel = new FileBrowserPanel();
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
		
		public FileBrowserPanel() {
			//TODO: Endre til gridbag layout
			setLayout(new GridLayout(2, 3));
			
			lblLayoutTemplate = new JLabel("Utseendemal: ");
			lblContentTemplate = new JLabel("Innholdsmal: ");
			
			fileChooser = new JFileChooser();
			fileChooser.setFileFilter(new FileNameExtensionFilter("OpenOffice", "odt"));
			
			txtContentTemplate = new JTextField();
			txtLayoutTemplate = new JTextField();
			
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
			int choice = fileChooser.showOpenDialog(FletteGUI.this);
			
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
		new FletteGUI();
	}
}
