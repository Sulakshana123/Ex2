/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.mavenproject1;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import javax.swing.DefaultListModel;
import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import javax.swing.ListModel;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Form extends javax.swing.JFrame {
    List <User> User = new ArrayList();
    private Component Jrame;

    public Form() {
        initComponents();
    }

  
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        namet = new javax.swing.JTextField();
        schoolt = new javax.swing.JTextField();
        aget = new javax.swing.JTextField();
        gendert = new javax.swing.JTextField();
        submit = new javax.swing.JButton();
        application = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        emailt = new javax.swing.JTextField();
        SP = new javax.swing.JScrollPane();
        SP1 = new javax.swing.JScrollPane();
        txt = new javax.swing.JTable();
        ET = new javax.swing.JButton();
        updatet = new javax.swing.JButton();
        deletet = new javax.swing.JButton();
        readt = new javax.swing.JButton();
        cleart = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setSize(new java.awt.Dimension(500, 500));

        jPanel1.setBackground(new java.awt.Color(0, 153, 153));

        namet.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                nametActionPerformed(evt);
            }
        });

        submit.setBackground(new java.awt.Color(204, 204, 255));
        submit.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        submit.setText("ADD");
        submit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                submitActionPerformed(evt);
            }
        });

        application.setFont(new java.awt.Font("Trebuchet MS", 3, 24)); // NOI18N
        application.setForeground(new java.awt.Color(255, 255, 255));
        application.setText("Application Form");

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Name");

        jLabel2.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("Email");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setText("School");

        jLabel4.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel4.setForeground(new java.awt.Color(255, 255, 255));
        jLabel4.setText("Age");

        jLabel5.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel5.setForeground(new java.awt.Color(255, 255, 255));
        jLabel5.setText("Gender");

        txt.setBackground(new java.awt.Color(204, 204, 255));
        txt.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        txt.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Name", "Email", "School", "Age", "Gender"
            }
        ));
        txt.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                txtMouseClicked(evt);
            }
        });
        SP1.setViewportView(txt);

        SP.setViewportView(SP1);

        ET.setBackground(new java.awt.Color(204, 204, 255));
        ET.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        ET.setText("Export Excel");
        ET.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ETActionPerformed(evt);
            }
        });

        updatet.setBackground(new java.awt.Color(204, 204, 255));
        updatet.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        updatet.setText("UPDATE");
        updatet.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                updatetActionPerformed(evt);
            }
        });

        deletet.setBackground(new java.awt.Color(204, 204, 255));
        deletet.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        deletet.setText("DELETE");
        deletet.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                deletetActionPerformed(evt);
            }
        });

        readt.setBackground(new java.awt.Color(204, 204, 255));
        readt.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        readt.setText("READ");
        readt.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                readtActionPerformed(evt);
            }
        });

        cleart.setBackground(new java.awt.Color(204, 204, 255));
        cleart.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        cleart.setText("CLEAR");
        cleart.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cleartActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(50, 50, 50)
                .addComponent(submit, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(28, 28, 28)
                .addComponent(updatet, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(31, 31, 31)
                .addComponent(deletet, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(ET, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(41, 41, 41)
                .addComponent(readt, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(40, 40, 40))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(application, javax.swing.GroupLayout.PREFERRED_SIZE, 310, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(61, 61, 61)
                .addComponent(cleart, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(32, 32, 32))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(79, 79, 79)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(SP, javax.swing.GroupLayout.PREFERRED_SIZE, 643, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(gendert, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(aget, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(schoolt, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(emailt, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(namet, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(170, 170, 170)))
                .addContainerGap(26, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(application, javax.swing.GroupLayout.PREFERRED_SIZE, 38, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(cleart))
                .addGap(28, 28, 28)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(namet, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(40, 40, 40)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(emailt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(34, 34, 34)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(schoolt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(34, 34, 34)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(aget, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(38, 38, 38)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(gendert, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(69, 69, 69)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(submit)
                    .addComponent(updatet)
                    .addComponent(deletet)
                    .addComponent(ET)
                    .addComponent(readt))
                .addGap(47, 47, 47)
                .addComponent(SP, javax.swing.GroupLayout.PREFERRED_SIZE, 183, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void nametActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_nametActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_nametActionPerformed
    
    
    
    private void submitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_submitActionPerformed
        submit(); 
        addRowTable();  
    }//GEN-LAST:event_submitActionPerformed

    private void ETActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ETActionPerformed
        addRowExcel();    
    }//GEN-LAST:event_ETActionPerformed

    private void updatetActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_updatetActionPerformed
        updateForm();
    }//GEN-LAST:event_updatetActionPerformed

    private void txtMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_txtMouseClicked
        txtmouse();
    }//GEN-LAST:event_txtMouseClicked

    private void deletetActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deletetActionPerformed
        dltForm();
    }//GEN-LAST:event_deletetActionPerformed

    private void readtActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_readtActionPerformed
        readExcel();
    }//GEN-LAST:event_readtActionPerformed

    private void cleartActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cleartActionPerformed
        clearText();
    }//GEN-LAST:event_cleartActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Form().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ET;
    private javax.swing.JScrollPane SP;
    private javax.swing.JScrollPane SP1;
    private javax.swing.JTextField aget;
    private javax.swing.JLabel application;
    private javax.swing.JButton cleart;
    private javax.swing.JButton deletet;
    private javax.swing.JTextField emailt;
    private javax.swing.JTextField gendert;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField namet;
    private javax.swing.JButton readt;
    private javax.swing.JTextField schoolt;
    private javax.swing.JButton submit;
    private javax.swing.JTable txt;
    private javax.swing.JButton updatet;
    // End of variables declaration//GEN-END:variables

    private void submit() {
        
       User Obj1 = new User();
       Obj1.setName(namet.getText());
       Obj1.setEmail(emailt.getText());
       Obj1.setSchool(schoolt.getText());
       Obj1.setAge(Integer.parseInt(aget.getText()));
       Obj1.setGender(gendert.getText());
       
       User.add(Obj1);
      
       
      
    }
    private void addRowTable() {
        DefaultTableModel mo = (DefaultTableModel)txt.getModel();
     
         Object rowData[] = new Object[5];
           for(int i = 0; i < User.size(); i++){
               
               rowData[0] = User.get(i).getName();
               rowData[1] = User.get(i).getEmail();
               rowData[2] = User.get(i).getSchool();
               rowData[3] = User.get(i).getAge();
               rowData[4] = User.get(i).getGender();
               mo.addRow(rowData);
           
           }
    
    }

    XSSFWorkbook workbook = new XSSFWorkbook();
    private void addRowExcel() {
            
       
       
        XSSFSheet sheet = workbook.createSheet();

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        for(int i =0; i < User.size(); i++){
            System.out.println(i);
            System.out.println(User.get(i).getName());
            System.err.println(User);
            System.out.println(User.get(0).toString());
            data.put(Integer.toString(i), new Object[] {User.get(i).getName(),User.get(i).getEmail(),User.get(i).getSchool(),User.get(i).getAge(),User.get(i).getGender()});
        }
          
        
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("Userdetails.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Users.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        }

    private void txtmouse() {
       DefaultTableModel mo = (DefaultTableModel) txt.getModel();
       int selectedRowIndex = txt.getSelectedRow();
        
         int index = txt.getSelectedRow();
         namet.setText(mo.getValueAt(index,0).toString());
         emailt.setText(mo.getValueAt(index,1).toString());
         schoolt.setText(mo.getValueAt(index,2).toString());
         aget.setText(mo.getValueAt(index,3).toString());
         gendert.setText(mo.getValueAt(index,4).toString());
    }

    private void updateForm() {
        int u = txt.getSelectedRow();
        DefaultTableModel mo = (DefaultTableModel)txt.getModel();
        if(u >= 0)
        {
            mo.setValueAt(namet.getText(), u, 0);
            mo.setValueAt(emailt.getText(), u, 1);
            mo.setValueAt(schoolt.getText(), u, 2);
            mo.setValueAt(aget.getText(), u, 3);
            mo.setValueAt(gendert.getText(), u, 4);
           
            JOptionPane.showMessageDialog(Jrame,"Update Successfully");
        }else{
            JOptionPane.showMessageDialog(Jrame,"Update UnSuccessfully");
        }
        
    }

    private void dltForm() {
         DefaultTableModel mo = (DefaultTableModel)txt.getModel();
       
            if(txt.getSelectedRow() != -1) 
            {
               // remove the selected row from the table model
               mo.removeRow(txt.getSelectedRow());
               JOptionPane.showMessageDialog(null, "Deleted successfully");

            }
    }

    private void readExcel() {
          try
        {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\User\\Desktop\\Mexxxar\\Ex2-main\\Userdetails.xlsx"));
 
           
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\r");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\r");
                            break;

                    }
                }
                System.out.println("");
            }
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
   
    }
    
    private void clearText(){
        namet.setText("");
        emailt.setText("");
        schoolt.setText("");
        aget.setText("");
        gendert.setText("");
    }
}
         

   

        

