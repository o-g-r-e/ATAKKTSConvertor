package com.yuriy.abc;

import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Observable;
import java.util.Observer;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.FileFilter;

public class MainWindow implements Observer
{	
	private File srcFile;
	private JLabel chooseFileLabel;
	private JTextField fileTextField;
	private JButton chooseFileButton;
	//private JProgressBar progressBar;
	private JTextArea progressTextArea;
	private JScrollPane progressScrollPane;
	private JButton startButton;
	private JFrame frame;
	private WindowEventsListener listener;
	
	private class ExcelFilesFilter extends FileFilter {

		@Override
		public boolean accept(File file) {
			
			return file.getAbsolutePath().endsWith(".xls") || file.getAbsolutePath().endsWith(".xlsx") || file.isDirectory();
		}

		@Override
		public String getDescription() {
			
			return "Excel documents (*.xls, *.xlsx)";
		}
	}
	
	static class ViewCommand {
		
		int line;
		private String data;
		
		public ViewCommand(int line, String data) {
			this.line = line;
			this.data = data;
		}

		public int getLine() {
			return line;
		}

		public void setLine(int line) {
			this.line = line;
		}

		public String getData() {
			return data;
		}

		public void setData(String data) {
			this.data = data;
		}
	}
	
	interface WindowEventsListener
	{
		public void onStartButtonClick();
		public void onFlieLoaded(String fileName);
	}
	
	public void addListener(WindowEventsListener listener)
	{
		this.listener = listener;
	}
	
	public MainWindow()
	{
		frame = new JFrame("Отчет по проверкам КТС АТАК");
		frame.setBounds(100,100,500,250);
		frame.setResizable(false);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		Container container = frame.getContentPane();
		
		container.setLayout(null);
		
		chooseFileLabel = new JLabel("Выберите файл:");
		chooseFileLabel.setBounds(101,25,150,30);
		
		container.add(chooseFileLabel);
		
		fileTextField = new JTextField(65535);
		fileTextField.setBounds(101,50,260,20);
		
		container.add(fileTextField);
		
		progressTextArea = new JTextArea(0,0);
		progressTextArea.setEditable(false);
		progressScrollPane = new JScrollPane(progressTextArea);
		progressScrollPane.setBounds(0,150,495,123);
		progressScrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		progressScrollPane.setVisible(false);
		/*progressBar = new JProgressBar();
		progressBar.setStringPainted(true);
		progressBar.setMinimum(0);
		progressBar.setMaximum(100);
		progressBar.setValue(0);
		progressBar.setBounds(10,180,470,15);*/
		container.add(progressScrollPane);
		
		startButton = createButton("Начать", 160,100,150,25, false);
		startButton.setEnabled(false);
		
		startButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if(srcFile != null)
                {
                	startProcess();
                	listener.onStartButtonClick();
                }
            }
        });
		
		container.add(startButton);
		
		chooseFileButton = createButton("...", 360, 50, 19, 19, true);
		
		chooseFileButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser(new File("."));
                fileChooser.setFileFilter(new ExcelFilesFilter());
                int ret = fileChooser.showDialog(null, "Открыть файл");                
                if (ret == JFileChooser.APPROVE_OPTION) {
                    fileLoaded(fileChooser.getSelectedFile());
                    listener.onFlieLoaded(srcFile.getAbsolutePath());
                }
            }
        });
		
		container.add(chooseFileButton);
		
		frame.setVisible(true);
	}
	
	private void fileLoaded(File file)
	{
		srcFile = file;
		fileTextField.setText(srcFile.getAbsolutePath());
        startButton.setEnabled(true);
	}
	
	private void startProcess()
	{
		frame.setBounds(100,100,500,300);
    	fileTextField.setText("");
    	fileTextField.setEnabled(false);
    	progressScrollPane.setVisible(true);
    	srcFile = null;
    	startButton.setEnabled(false);
    	chooseFileButton.setEnabled(false);
	}
	
	private JButton createButton(String title, int x, int y, int width, int height, boolean enable)
	{
		JButton button = new JButton(title);
		button.setBounds(x, y, width, height);
		button.setEnabled(enable);
		
		return button;
	}
	
	public void showException(String exceptionMessage)
	{
		JOptionPane.showMessageDialog(null, exceptionMessage, "Exception", JOptionPane.ERROR_MESSAGE);
	}

	@Override
	public void update(Observable arg0, Object message) {
		
		ViewCommand viewCommand = (ViewCommand)message;
		
		SwingUtilities.invokeLater(new Runnable() {
            public void run() {
            	
            	String progressText = progressTextArea.getText();
            	
            	String[] lines = progressText.split("\n");
        		
        		if(lines.length == 1 && lines[0].equals(""))
        		{
        			progressText = viewCommand.getData()+"\n";
        		}
        		else
        		{
            		if(lines.length > viewCommand.getLine())
            		{
            			lines[viewCommand.getLine()] = viewCommand.getData();
            			progressText = "";
            			for (int i = 0; i < lines.length; i++) {
                			progressText += lines[i]+(i<lines.length-1?"\n":"");
    					}
            		}
            		else
            		{
            			progressText += "\n"+viewCommand.getData();
            		}
        		}
            	
            	progressTextArea.setText(progressText);
            	//progressTextArea.setText(sMessage);
            }
        });
	}
}