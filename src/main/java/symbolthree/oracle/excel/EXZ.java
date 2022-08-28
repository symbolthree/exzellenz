/******************************************************************************
 *
 * ≡ EXZELLENZ ≡
 * Copyright (C) 2009-2022 Christopher Ho 
 * All Rights Reserved, http://www.symbolthree.com
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 *
 * E-mail: christopher.ho@symbolthree.com
 *
******************************************************************************/

package symbolthree.oracle.excel;

//~--- JDK imports ------------------------------------------------------------

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.*;
import java.net.URI;
import java.util.Enumeration;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

public class EXZ implements Constants, ActionListener {
    private JFrame          frame = new JFrame();
    //private final JTextArea text  = new JTextArea();
    private final JTextPane text      = new JTextPane();
    private HTMLEditorKit   kit       = new HTMLEditorKit();
    private HTMLDocument    styleDoc  = new HTMLDocument();
    private JPopupMenu      popupMenu = new JPopupMenu();
    private String          inputFileName;
    
    private JMenuItem    openMenu;
    private JMenuItem    clearMenu;
    private JMenu        logMenu;
    private JMenu        logLevelMenu;
    private JMenu        logIntervalMenu;
    private JRadioButton logDebugMenu;    
    private JRadioButton logInfoMenu;
    private JRadioButton logWarnMenu;
    private JMenuItem    helpMenu;
    private JMenuItem    aboutMenu;
    private JMenuItem    exitMenu;
    //private JCheckBoxMenuItem newFileMenu;
    private JFileChooser fc = new JFileChooser();
    private ButtonGroup logLevelGroup    = new ButtonGroup();
    private ButtonGroup logIntervalGroup = new ButtonGroup();
    
    private EXZELLENZ exzellenz = null;  
    private Thread    thread    = null;
    
    PipedInputStream        piErr;
    PipedInputStream        piOut;
    PipedOutputStream       poErr;
    PipedOutputStream       poOut;

    public EXZ() throws Exception {
    	showSplashScreen();
        System.setProperty(RUN_MODE, RUNMODE_GUI);
        EXZHelper.initializeLogging();
        EXZHelper.getVersion();
        redirectSystemStreams();
        createPopupMenu();
        text.putClientProperty(JEditorPane.HONOR_DISPLAY_PROPERTIES, true);        
        text.setComponentPopupMenu(popupMenu);
        text.setEditorKit(kit);
        text.setDocument(styleDoc);

        addText("<font color='" + LOG_INFO_COLOR + "'>" + EXZI18N.inst().get("MSG.INSTRUCTION") + "</font><br/></br/>");
    }

    public static void main(String[] args) {
        try {
            EXZ exz = new EXZ();

            exz.start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void start() throws Exception {
        frame.setTitle("EXZELLENZ " + System.getProperty(EXZELLENZ_VERSION));

        BufferedImage icon = null;

        try {
            icon = ImageIO.read(frame.getClass().getResource("/symbolthree/oracle/excel/exzellenz_icon.gif"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        frame.setIconImage(icon);

        Class<?> clazz  = Class.forName("java.awt.Font");
        Field field     = clazz.getField(EXZProp.instance().getStr("FONT_STYLE"));
        int   fontStyle = field.getInt(EXZProp.instance().getStr("FONT_STYLE"));
        Font  font      = new Font(EXZProp.instance().getStr("FONT_NAME"), fontStyle,
                                   EXZProp.instance().getInt("FONT_SIZE"));
        text.setFont(font);

        Color fontColor = getColor(EXZProp.instance().getStr("FONT_COLOR"), Color.WHITE);
        text.setForeground(fontColor);
        
        Color bgColor = getColor(EXZProp.instance().getStr("BACKGROUND_COLOR"), Color.BLACK);
        text.setBackground(bgColor);
        
        //text.setText(EXZI18N.inst().get("MSG.INSTRUCTION") + "\n\n");
        text.setEditable(false);
        frame.getContentPane().add(new JScrollPane(text), BorderLayout.CENTER);

        Dimension screenSize   = Toolkit.getDefaultToolkit().getScreenSize();
        int       windowWidth  = EXZProp.instance().getInt("WINDOW_WIDTH");
        int       windowHeight = EXZProp.instance().getInt("WINDOW_HEIGHT");

        frame.setSize(windowWidth, windowHeight);
        int xPos = (int) (screenSize.getWidth() - windowWidth) / 2;
        int yPos = (int) (screenSize.getHeight() - windowHeight) / 2;
        if (xPos < 0) xPos = 0;
        if (yPos < 0) yPos = 0;
        frame.setLocation(xPos, yPos);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
        
        frame.addWindowListener(new WindowAdapter() {
           public void windowClosing(WindowEvent e) {
             EXZProp.instance().saveSettings();
             System.exit(0);
           }
        });
        
        frame.addComponentListener(new ComponentAdapter() {
        	public void componentResized(ComponentEvent evt) {
        	  EXZProp.instance().setStr(WINDOW_WIDTH, String.valueOf(frame.getWidth()));
        	  EXZProp.instance().setStr(WINDOW_HEIGHT, String.valueOf(frame.getHeight()));
            }
        });
        
        new FileDrop(System.out, text, new FileDrop.Listener() {
            public void filesDropped(File[] files) {
                boolean startProcess = true;

                if (files.length > 1) {
                    addText(EXZI18N.inst().get("ERR.MULTI_FILE"));
                    startProcess = false;
                } else {
                    inputFileName = files[0].getAbsolutePath();

                    String fileExtension = EXZHelper.getExtension(files[0]);

                    if (!fileExtension.equalsIgnoreCase("XLS") &&
                        !fileExtension.equalsIgnoreCase("XLSX")) {
                        addText(EXZI18N.inst().get("ERR.INVALID_EXT", inputFileName));
                        startProcess = false;
                    }
                }

                if (startProcess) {
                  runEXZELLENZ(inputFileName);
                }
            }    // end filesDropped
        });    // end FileDrop.Listener
    }    // end main

    private void runEXZELLENZ(final String _file) {
      EXZProp.instance().setStr(LAST_FILE, _file);
      EXZProp.instance().saveSettings();
      
      addText(EXZI18N.inst().get("MSG.PROCESSING", _file));

      exzellenz = new EXZELLENZ();
      exzellenz.setFile(_file);
      thread = new Thread(exzellenz);
      thread.start();
    }
    
    private void redirectSystemStreams() {
        OutputStream out = new OutputStream() {
            public void write(int b) throws IOException {
                addText(String.valueOf((char) b));
            }
            public void write(byte[] b, int off, int len) throws IOException {
              addText(new String(b, off, len));
            }
            public void write(byte[] b) throws IOException {
                write(b, 0, b.length);
            }
        };

        System.setOut(new PrintStream(out, true));
        System.setErr(new PrintStream(out, true));
    }
    
    private void createPopupMenu() {
      openMenu  = new JMenuItem(EXZI18N.inst().get("MENU.OPEN"));
      clearMenu = new JMenuItem(EXZI18N.inst().get("MENU.CLEAR"));
      /*
      if (EXZProp.instance().getBoolean(SAVE_NEW_FILE)) {
        newFileMenu = new JCheckBoxMenuItem(EXZI18N.inst().get("MENU.NEWFILE"), true);        
      } else {
        newFileMenu = new JCheckBoxMenuItem(EXZI18N.inst().get("MENU.NEWFILE"), false);        
      }
      */
      logMenu  = new JMenu(EXZI18N.inst().get("MENU.LOGGING"));
      logLevelMenu = new JMenu(EXZI18N.inst().get("MENU.LOG_LEVEL"));     
      logWarnMenu = new JRadioButton(EXZI18N.inst().get("MENU.LOG_WARN"));      
      logInfoMenu = new JRadioButton(EXZI18N.inst().get("MENU.LOG_INFO"));
      logDebugMenu = new JRadioButton(EXZI18N.inst().get("MENU.LOG_DEBUG"));
      logWarnMenu.addActionListener(this);
      logInfoMenu.addActionListener(this);
      logDebugMenu.addActionListener(this);
      
      logLevelGroup.add(logWarnMenu);
      logLevelGroup.add(logInfoMenu);
      logLevelGroup.add(logDebugMenu);
      logLevelMenu.add(logWarnMenu);
      logLevelMenu.add(logInfoMenu);
      logLevelMenu.add(logDebugMenu);
      
      logIntervalMenu = new JMenu(EXZI18N.inst().get("MENU.LOG_INTERVAL"));
      JRadioButton log10Lines   = new JRadioButton("10");
      JRadioButton log50Lines   = new JRadioButton("50");
      JRadioButton log100Lines  = new JRadioButton("100");
      JRadioButton log500Lines  = new JRadioButton("500");
      JRadioButton log1000Lines = new JRadioButton("1000");
      log10Lines.addActionListener(this);
      log50Lines.addActionListener(this);
      log100Lines.addActionListener(this);
      log500Lines.addActionListener(this);
      log1000Lines.addActionListener(this);
      
      logIntervalGroup.add(log10Lines);
      logIntervalGroup.add(log50Lines);
      logIntervalGroup.add(log100Lines);
      logIntervalGroup.add(log500Lines);
      logIntervalGroup.add(log1000Lines);
      
      logIntervalMenu.add(log10Lines);
      logIntervalMenu.add(log50Lines);
      logIntervalMenu.add(log100Lines);
      logIntervalMenu.add(log500Lines);
      logIntervalMenu.add(log1000Lines);
      logMenu.add(logLevelMenu);
      logMenu.add(logIntervalMenu);

      setLogging();
      
      helpMenu  = new JMenuItem(EXZI18N.inst().get("MENU.HELP"));
      aboutMenu = new JMenuItem(EXZI18N.inst().get("MENU.ABOUT"));      
      exitMenu = new JMenuItem(EXZI18N.inst().get("MENU.EXIT"));
      openMenu.addActionListener(this);
      clearMenu.addActionListener(this);
      helpMenu.addActionListener(this);
      aboutMenu.addActionListener(this);
      exitMenu.addActionListener(this);
      //newFileMenu.addActionListener(this);
      popupMenu.add(openMenu);
      //popupMenu.add(newFileMenu);
      popupMenu.add(clearMenu);
      popupMenu.add(logMenu);
      popupMenu.addSeparator();
      popupMenu.add(helpMenu);
      popupMenu.add(aboutMenu);
      popupMenu.add(exitMenu);      
      
      File lastFile = new File(EXZProp.instance().getStr(LAST_FILE));
      if (lastFile != null) fc.setCurrentDirectory(lastFile);
      fc.addChoosableFileFilter(new FileFilter() {
        public boolean accept(File f) {
            if (f.isDirectory()) {
                return true;
            }

            String extension = EXZHelper.getExtension(f);

            if (extension != null) {
                if (extension.equalsIgnoreCase("XLS") || 
                	extension.equalsIgnoreCase("XLSX") ||
                	extension.equalsIgnoreCase("XLSB")) {
                    return true;
                } else {
                    return false;
                }
            } else {
                return false;
            }
        }
        public String getDescription() {
            return "Microsoft Excel file only";
        }
    });      
    }

    public void actionPerformed(ActionEvent e) {
      if (e.getSource()==openMenu) {
        int returnVal = fc.showOpenDialog(frame);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
          File file = fc.getSelectedFile();
          runEXZELLENZ(file.getAbsolutePath());
        }
        
      } else if (e.getSource()==clearMenu) {
        text.setText("");
      
        /*
      } else if (e.getSource()==newFileMenu) {
        if (EXZProp.instance().getBoolean(SAVE_NEW_FILE)) {
          newFileMenu.setSelected(false);
          EXZProp.instance().setStr(SAVE_NEW_FILE, "FALSE");
        } else {
          newFileMenu.setSelected(true);
          EXZProp.instance().setStr(SAVE_NEW_FILE, "TRUE");          
        }
        */
      
      } else if (e.getSource()==helpMenu) {
    	  String helpLink = EXZI18N.inst().get("MENU.HELP_LINK");
     	  try {
	    	  if (Desktop.isDesktopSupported()) {
	    		 if (helpLink.toLowerCase().startsWith("http")) {
	    	       Desktop.getDesktop().browse(new URI(helpLink));
	    		 } else {
	    	       Desktop.getDesktop().open(new File(helpLink));
	    		 }
	    	  }
    	  } catch (Exception ex) {
    		  addText("Unable to open help file or link");
    	  }
        
      } else if (e.getSource()==aboutMenu) {
          AboutDialog f =  new AboutDialog();
          f.setLocationRelativeTo(frame);
          f.setVisible(true);
          
      } else if (e.getSource()==exitMenu) {
        EXZProp.instance().saveSettings();
        System.exit(0); 
      } else if (e.getSource() instanceof JRadioButton) {
    	  setLogging((JRadioButton)e.getSource());
      }
    }

    public void addText(final String str) {
      SwingUtilities.invokeLater(new Runnable() {
          public void run() {
        	try {
			  if (str.startsWith("$$$")) {
				  String _str = str.substring(3);
				  // TODO pause thread
			      int response = JOptionPane.showConfirmDialog(frame, _str, "Confirm",
				            JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
			      if (response==JOptionPane.YES_OPTION) {
			    	 System.setProperty(CONFIRM_RESPONSE, "Y"); 
			      } else {
			    	 System.setProperty(CONFIRM_RESPONSE, "N");
			      }
			  } else {
			    kit.insertHTML(styleDoc, styleDoc.getLength(), str, 0, 0, null);
			  }
			  
			} catch (BadLocationException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
          }
      });
    }
    
    
    private void setLogging(JRadioButton button) {
    	if (button==logDebugMenu && logDebugMenu.isSelected()) {  
    		EXZProp.instance().setStr(EXZ_LOG_LEVEL, "DEBUG");  
    	} else if (button==logInfoMenu && logInfoMenu.isSelected()) {  
    		EXZProp.instance().setStr(EXZ_LOG_LEVEL, "INFO");
        } else if (button==logWarnMenu && logWarnMenu.isSelected()) {
    		EXZProp.instance().setStr(EXZ_LOG_LEVEL, "WARN");
        } else {
          String logInterval = button.getText();  	
          EXZProp.instance().setStr(EXZ_LOG_INTERVAL, logInterval);
        }
    }
    
    private void setLogging() {
    	String logLevel = EXZProp.instance().getStr(EXZ_LOG_LEVEL);
        if (logLevel.equals("WARN")) {
        	logWarnMenu.setSelected(true);
        } else if (logLevel.equals("INFO")) {
        	logInfoMenu.setSelected(true);
        } else if (logLevel.equals("DEBUG")) {
        	logDebugMenu.setSelected(true);
        }
        String logInt = EXZProp.instance().getStr(EXZ_LOG_INTERVAL);
        Enumeration<AbstractButton> em = logIntervalGroup.getElements(); 
        while (em.hasMoreElements()) {
        	JRadioButton button = (JRadioButton)em.nextElement();
        	String label = button.getText();
        	if (logInt.equals(label)) {
        		button.setSelected(true);
        	}
        }
        
    }
    
    private void showSplashScreen() {
        SplashScreen splash = SplashScreen.getSplashScreen();
        if (splash == null) return;
        Graphics2D g = splash.createGraphics();
        if (g == null) return;
        try {
            Thread.sleep(500);
        }
        catch(InterruptedException e) {}       
    }
    
    private Color getColor(String color, Color defaultValue) {
        if (color == null) {
            return defaultValue;
        }

        try {
            return Color.decode(color.toUpperCase());
        } catch (NumberFormatException nfe) {
            try {
                final Field f = Color.class.getField(color.toUpperCase());

                return (Color) f.get(null);
            } catch (Exception ce) {
                return defaultValue;
            }
        }
    }    
    
}
