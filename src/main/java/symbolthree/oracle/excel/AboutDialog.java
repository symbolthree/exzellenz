/******************************************************************************
 *
 * ≡ EXZELLENZ ≡
 * Copyright (C) 2009-2016 Christopher Ho 
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
 * ================================================
 *
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/AboutDialog.java $
 * $Author: Christopher Ho $
 * $Date: 8/02/16 11:22p $
 * $Revision: 4 $
******************************************************************************/

package symbolthree.oracle.excel;

import java.awt.Component;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Properties;
import java.util.jar.Attributes;
import java.util.jar.JarFile;
import java.util.jar.Manifest;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTable;
import javax.swing.RowSorter;
import javax.swing.RowSorter.SortKey;
import javax.swing.SortOrder;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;

import oracle.jdbc.OracleDriver;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.Namespace;
import org.jdom2.filter.Filters;
import org.jdom2.input.SAXBuilder;
import org.jdom2.xpath.XPathExpression;
import org.jdom2.xpath.XPathFactory;

public class AboutDialog extends JDialog implements ActionListener, Constants {

	private static final long serialVersionUID = 8564475892772711012L;
	private Properties versions = new Properties();
	private JButton closeBTN    = new JButton(EXZI18N.inst().get("MENU.CLOSE"));	
	
	public AboutDialog() {
		
		initVersions();
		
		this.setTitle(EXZI18N.inst().get("MENU.ABOUT"));
        //this.setSize(450, 300);
        //BoxLayout boxLayout = new BoxLayout(this, BoxLayout.Y_AXIS);
        this.setLayout(new GridBagLayout());
    	GridBagConstraints GC = new GridBagConstraints();
    	
		JLabel label1  = new JLabel("≡ EXZELLENZ ≡");
		label1.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 14));
		JLabel version = new JLabel(EXZHelper.getVersionWithTimestamp());
		version.setFont(new Font(Font.DIALOG, Font.ITALIC, 12));
		JLabel label2 = new JLabel(EXZHelper.getAuthorLine());
		JLabel label3 = new JLabel("<html><a href='#'>Home Page</a></html>");
		label3.setCursor(new Cursor(Cursor.HAND_CURSOR));
		label3.addMouseListener(new MouseAdapter() {
			@Override
            public void mouseClicked(MouseEvent e) {
			  try {
				Desktop.getDesktop().browse(new URI(EXZI18N.inst().get("MENU.ABOUT_LINK")));
              } catch (Exception ex) {
			  }
			}
		});
		JSeparator sep = new JSeparator();
		
		GC.gridx=0;
		GC.gridy=0;
		this.add(label1, GC);
		GC.gridy=1;
		this.add(version, GC);
		GC.gridy=2;
		this.add(label2, GC);
		GC.gridy=3;
		this.add(label3, GC);
		GC.gridy=4;
		this.add(sep, GC);
        
        JScrollPane jsp = new JScrollPane();
        JTable table = new JTable();
        String[] columnNames = {"Component", "Version"};
        
        DefaultTableModel tableModel = new DefaultTableModel(columnNames, 0) {
			private static final long serialVersionUID = 4886287779669939039L;
			@Override
        	   public boolean isCellEditable(int row, int column) {
        	       return false;
        	   }
        };
        
        table.setModel(tableModel);
        Enumeration<?> enumKey = versions.keys();
        while (enumKey.hasMoreElements()) {
          String key = (String)enumKey.nextElement();
          String val = versions.getProperty(key);
          tableModel.addRow(new Object[]{key, val});
        }
        
        table.setAutoCreateRowSorter(true);
        RowSorter<? extends TableModel> sorter = table.getRowSorter();
        List<SortKey> sortKeys = new ArrayList<SortKey>();
        sortKeys.add(new RowSorter.SortKey(0, SortOrder.ASCENDING));
        sorter.setSortKeys(sortKeys);

        JPanel    panel     = new JPanel();        
        table.setFillsViewportHeight(true);
        jsp.setViewportView(table);
        panel.add(jsp); 
        GC.gridy=5;
        this.add(panel, GC);
        
        closeBTN.addActionListener(this);
        closeBTN.setAlignmentX(Component.CENTER_ALIGNMENT);
        GC.gridy=6;
        this.add(closeBTN, GC);

        this.pack();
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		if (e.getSource()==closeBTN) {
        	this.setVisible(false);			
		}
		
	}
	
	private void initVersions() {
	  try {
        
        addVersion("Oracle JDBC Driver", OracleDriver.getDriverVersion());
        
        addVersion("JRE Version", System.getProperty("java.version") + " " + System.getProperty("sun.arch.data.model") + " bit");

        addVersion("Java Home", System.getProperty("java.home"));
        
        //readManifest();
        readPOM();
      
        /*
        File fndextJar = new File(System.getProperty("user.dir") + File.separator + "lib", "fndext.jar");
        ZipFile zipFile = new ZipFile(fndextJar);
        String comment = zipFile.getComment();
        String ver = comment.substring(comment.indexOf("fndext.jar")+10).trim();
        addVersion("fndext.jar", ver);
        zipFile.close();
         */
        
		} catch (Exception e) {
			e.printStackTrace();
		}
	  
	}
	
	private void addVersion(String key, String val) {
		if (val==null) val="";
		versions.setProperty(key, val);
	}
	
	private void readPOM() throws Exception {
		Namespace   ns = Namespace.getNamespace("ns", "http://maven.apache.org/POM/4.0.0");
		InputStream is = this.getClass().getResourceAsStream("/META-INF/maven/symbolthree.oracle.excel/exzellenz/pom.xml");
		
		SAXBuilder jdomBuilder = new SAXBuilder();
		Document jdomDoc = jdomBuilder.build(is);
		
		XPathFactory xpfac = XPathFactory.instance();
		XPathExpression<Element> xp = xpfac.compile("//ns:poi.version", Filters.element(), null, ns);
	    Element ele = xp.evaluateFirst(jdomDoc);
	    addVersion("org.apache.poi", ele.getValue());
	    
	    xp = xpfac.compile("//ns:dependency", Filters.element(), null, ns);
	    List<Element> eles = xp.evaluate(jdomDoc);
	    for (Element e : eles) {
	    	String groupId = e.getChild("groupId", ns).getText();	    	
	    	String version = e.getChild("version", ns).getText();
	    	if (! groupId.startsWith("oracle") && groupId.indexOf("poi") < 0) {
	    		addVersion(groupId, version);
	    	} 
	    }	    
    }
	
	private void readManifest() throws Exception {
	    Enumeration<?> resEnum;
        resEnum = this.getClass().getClassLoader().getResources(JarFile.MANIFEST_NAME);
        while (resEnum.hasMoreElements()) {
	        URL url = (URL)resEnum.nextElement();
	        InputStream is = url.openStream();
	        if (is != null) {
	            Manifest manifest = new Manifest(is);
	            Attributes mainAttribs = manifest.getMainAttributes();
	            String title = mainAttribs.getValue("Implementation-Title");
	            String version = mainAttribs.getValue("Implementation-Version");
            
	            if (title!=null) {
	              if(title.startsWith("org.dom4j")) {
	            	  addVersion(title, version);
	              }
	              if (title.startsWith("Apache POI")) {
	            	  addVersion(title, version);
	              }
	              if (title.indexOf("xmlbeans") > 0) {
	            	  addVersion(title, version);
	              }
	              if (title.startsWith("Commons IO")) {
	            	  addVersion(title, version);
	              }
	              if (title.startsWith("StAX")) {
	            	  addVersion(title, version);
	              }	              
	              if (title.indexOf("i18n") > 0) {
	            	  addVersion(title, version);
	              }	              
	            }
	        }
        }
	}
}