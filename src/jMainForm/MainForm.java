/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jMainForm;

import java.awt.Color;
import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;

import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Arrays;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.SwingConstants;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.InetAddress;
import java.net.ProtocolException;
import java.net.UnknownHostException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.ListIterator;
import java.util.logging.Level;
import java.util.logging.Logger;


/**
 *
 * @author earmenda
 */

class PassWordDialog extends JDialog {

    private final JLabel jlblUsername = new JLabel("Usuario");
    private final JLabel jlblPassword = new JLabel("Password");

    private final JTextField jtfUsername = new JTextField(15);
    private final JPasswordField jpfPassword = new JPasswordField();

    private final JButton jbtOk = new JButton("Login");
    private final JButton jbtCancel = new JButton("Cancelar");

    private final JLabel jlblStatus = new JLabel(" ");

    public PassWordDialog() {
        this(null, true);
    }

    public PassWordDialog(final JFrame parent, boolean modal) {
        super(parent, modal);

        JPanel p3 = new JPanel(new GridLayout(2, 1));
        p3.add(jlblUsername);
        p3.add(jlblPassword);

        JPanel p4 = new JPanel(new GridLayout(2, 1));
        p4.add(jtfUsername);
        p4.add(jpfPassword);

        JPanel p1 = new JPanel();
        p1.add(p3);
        p1.add(p4);

        JPanel p2 = new JPanel();
        p2.add(jbtOk);
        p2.add(jbtCancel);

        JPanel p5 = new JPanel(new BorderLayout());
        p5.add(p2, BorderLayout.CENTER);
        p5.add(jlblStatus, BorderLayout.NORTH);
        jlblStatus.setForeground(Color.RED);
        jlblStatus.setHorizontalAlignment(SwingConstants.CENTER);

        setLayout(new BorderLayout());
        add(p1, BorderLayout.CENTER);
        add(p5, BorderLayout.SOUTH);
        pack();
        setLocationRelativeTo(null);
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);

        addWindowListener(new WindowAdapter() {  
            @Override
            public void windowClosing(WindowEvent e) {  
                System.exit(0);  
            }  
        });


        jbtOk.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (Arrays.equals("06122013".toCharArray(), jpfPassword.getPassword())
                        && "capacitacion".equals(jtfUsername.getText())) {
                    parent.setVisible(true);
                    setVisible(false);
                } else {
                    jlblStatus.setText("Usuario o contraseña incorrectas");
                }
            }
        });
        
        jbtCancel.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                setVisible(false);
                parent.dispose();
                System.exit(0);
            }
        });
    }
}

public class MainForm extends javax.swing.JFrame {

    /**
     * Creates new form MainForm
     */
    private PassWordDialog passDialog;
    
    public MainForm() throws MalformedURLException, IOException {
        
        initComponents();
        passDialog = new PassWordDialog(this, true);
        passDialog.setVisible(true);                        
        logAccess();
        /*
        
        //Testing for new diplomas
        String[] args = {"/Users/earmenda/Desktop/DirectorioBAE.xlsx","/Users/earmenda/Desktop/test","/Users/earmenda/Desktop/FormatoCertificado.xlsx","TSI. Jorge Antonio Razón Gil","Manuel Anguiano Razón","false","false","false","true","true","true","false","false","false","false","","/Users/earmenda/Google Drive/Visio-Expressway.pdf"};
        try{
            cdiscisa.Cdiscisa.main(args);
        }catch(Exception ex){
            System.out.println(ex.toString());
        }
        
        System.exit(0);
       */
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        chkDiplomas = new javax.swing.JCheckBox();
        lblCheckDocumentos = new javax.swing.JLabel();
        chkDiplomaFirma = new javax.swing.JCheckBox();
        chkDC3 = new javax.swing.JCheckBox();
        chkConstanciasFirma1 = new javax.swing.JCheckBox();
        chkDiplomasLogo = new javax.swing.JCheckBox();
        chkConstanciasLogo = new javax.swing.JCheckBox();
        chkDC3Logo = new javax.swing.JCheckBox();
        chkDC3Firma = new javax.swing.JCheckBox();
        chkConstancias = new javax.swing.JCheckBox();
        chkArchivoCompilado = new javax.swing.JCheckBox();
        lblInstructor = new javax.swing.JLabel();
        cboxUnidadCapacitadora = new javax.swing.JComboBox<>();
        cboxInstructor = new javax.swing.JComboBox<>();
        lblUnidadCapacitadora = new javax.swing.JLabel();
        generar = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        lblDirectorio = new javax.swing.JLabel();
        txtDirectorio = new javax.swing.JTextField();
        btnExplorarDirectorio = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        btnListaAuto = new javax.swing.JButton();
        txtListaAuto = new javax.swing.JTextField();
        lblListaAuto = new javax.swing.JLabel();
        lblLista = new javax.swing.JLabel();
        btnLista = new javax.swing.JButton();
        txtLista = new javax.swing.JTextField();
        jPanel4 = new javax.swing.JPanel();
        btnExplorarCarpetas = new javax.swing.JButton();
        txtGuardarDocs = new javax.swing.JTextField();
        lblDirectorioGuardar = new javax.swing.JLabel();
        jPanel5 = new javax.swing.JPanel();
        txtRegistro = new javax.swing.JTextField();
        btnExplorarRegistro = new javax.swing.JButton();
        chkRegistro = new javax.swing.JCheckBox();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("ISCISA - Generar Documentos");
        setBackground(new java.awt.Color(0, 51, 153));
        setLocation(new java.awt.Point(100, 200));
        setName("ISCISA Generar Documentos"); // NOI18N
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        jPanel1.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        chkDiplomas.setFont(new java.awt.Font("Helvetica", 3, 14)); // NOI18N
        chkDiplomas.setSelected(true);
        chkDiplomas.setText("Diplomas");
        chkDiplomas.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chkDiplomasActionPerformed(evt);
            }
        });

        lblCheckDocumentos.setFont(new java.awt.Font("Lucida Grande", 1, 14)); // NOI18N
        lblCheckDocumentos.setText("¿Que documentos se van a generar?");

        chkDiplomaFirma.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkDiplomaFirma.setSelected(true);
        chkDiplomaFirma.setText("¿Incluir firma del Instructor?");

        chkDC3.setFont(new java.awt.Font("Helvetica", 3, 14)); // NOI18N
        chkDC3.setSelected(true);
        chkDC3.setText("DC-3 (Alumnos Aprobados)");
        chkDC3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chkDC3ActionPerformed(evt);
            }
        });

        chkConstanciasFirma1.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkConstanciasFirma1.setSelected(true);
        chkConstanciasFirma1.setText("¿Incluir firma del Instructor?");

        chkDiplomasLogo.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkDiplomasLogo.setSelected(true);
        chkDiplomasLogo.setText("¿Incluir logo de ISCISA?");

        chkConstanciasLogo.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkConstanciasLogo.setSelected(true);
        chkConstanciasLogo.setText("¿Incluir logo de ISCISA?");

        chkDC3Logo.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkDC3Logo.setSelected(true);
        chkDC3Logo.setText("¿Incluir logo de ISCISA?");

        chkDC3Firma.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        chkDC3Firma.setSelected(true);
        chkDC3Firma.setText("¿Incluir firma del Instructor?");

        chkConstancias.setFont(new java.awt.Font("Helvetica", 3, 14)); // NOI18N
        chkConstancias.setSelected(true);
        chkConstancias.setText("Constancias");
        chkConstancias.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chkConstanciasActionPerformed(evt);
            }
        });

        chkArchivoCompilado.setFont(new java.awt.Font("Helvetica", 3, 14)); // NOI18N
        chkArchivoCompilado.setSelected(true);
        chkArchivoCompilado.setText("¿Generar un archivo único compilando los anteriores?");
        chkArchivoCompilado.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chkArchivoCompiladoActionPerformed(evt);
            }
        });

        lblInstructor.setFont(new java.awt.Font("Helvetica", 1, 14)); // NOI18N
        lblInstructor.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblInstructor.setText("Instructor");

        cboxUnidadCapacitadora.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        cboxUnidadCapacitadora.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "TSI. Jorge Antonio Razón Gil", "Grupo ISCISA S.A. de C.V." }));
        cboxUnidadCapacitadora.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cboxUnidadCapacitadoraActionPerformed(evt);
            }
        });

        cboxInstructor.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        cboxInstructor.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "TSI. Jorge Antonio Razón Gil" }));

        lblUnidadCapacitadora.setFont(new java.awt.Font("Helvetica", 1, 14)); // NOI18N
        lblUnidadCapacitadora.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblUnidadCapacitadora.setText("Unidad Capacitadora");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(1, 1, 1)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(chkConstancias)
                            .addComponent(chkDiplomas)
                            .addComponent(lblCheckDocumentos)))
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(chkDiplomasLogo)
                                .addComponent(chkConstanciasLogo))
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(jPanel1Layout.createSequentialGroup()
                                    .addGap(49, 49, 49)
                                    .addComponent(chkConstanciasFirma1))
                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(chkDiplomaFirma))))
                        .addGroup(jPanel1Layout.createSequentialGroup()
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(chkDC3)
                                .addGroup(jPanel1Layout.createSequentialGroup()
                                    .addGap(29, 29, 29)
                                    .addComponent(chkDC3Logo)))
                            .addGap(23, 23, 23)
                            .addComponent(chkDC3Firma)))
                    .addComponent(chkArchivoCompilado))
                .addContainerGap(21, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(cboxUnidadCapacitadora, 0, 241, Short.MAX_VALUE)
                    .addComponent(lblInstructor, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(cboxInstructor, javax.swing.GroupLayout.PREFERRED_SIZE, 239, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lblUnidadCapacitadora, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(100, 100, 100))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lblCheckDocumentos)
                .addGap(18, 18, 18)
                .addComponent(chkDiplomas)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(chkDiplomasLogo)
                    .addComponent(chkDiplomaFirma))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(chkConstancias)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(chkConstanciasLogo)
                    .addComponent(chkConstanciasFirma1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(chkDC3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(chkDC3Firma)
                    .addComponent(chkDC3Logo))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(chkArchivoCompilado)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 43, Short.MAX_VALUE)
                .addComponent(lblUnidadCapacitadora)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(cboxUnidadCapacitadora, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(lblInstructor)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(cboxInstructor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        generar.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        generar.setText("Generar Documentos");
        generar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                generarActionPerformed(evt);
            }
        });

        jPanel2.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        lblDirectorio.setFont(new java.awt.Font("Helvetica Neue", 1, 14)); // NOI18N
        lblDirectorio.setText("¿Donde se encuentra el archivo con las direcciones de las sucursales?");

        txtDirectorio.setFont(new java.awt.Font("Helvetica Neue", 2, 12)); // NOI18N
        txtDirectorio.setHorizontalAlignment(javax.swing.JTextField.LEFT);
        txtDirectorio.setText("Selecciona el archivo Excel  que continene el directorio...");
        txtDirectorio.setToolTipText("");
        txtDirectorio.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txtDirectorio.setDoubleBuffered(true);
        txtDirectorio.setEnabled(false);
        txtDirectorio.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtDirectorioFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                txtDirectorioFocusLost(evt);
            }
        });
        txtDirectorio.addInputMethodListener(new java.awt.event.InputMethodListener() {
            public void inputMethodTextChanged(java.awt.event.InputMethodEvent evt) {
                txtDirectorioInputMethodTextChanged(evt);
            }
            public void caretPositionChanged(java.awt.event.InputMethodEvent evt) {
            }
        });
        txtDirectorio.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtDirectorioActionPerformed(evt);
            }
        });
        txtDirectorio.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                txtDirectorioPropertyChange(evt);
            }
        });

        btnExplorarDirectorio.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        btnExplorarDirectorio.setText("Explorar Archivos");
        btnExplorarDirectorio.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExplorarDirectorioActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel2Layout.createSequentialGroup()
                        .addGap(17, 17, 17)
                        .addComponent(lblDirectorio)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel2Layout.createSequentialGroup()
                        .addComponent(txtDirectorio)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(btnExplorarDirectorio)))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lblDirectorio, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 22, Short.MAX_VALUE)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnExplorarDirectorio)
                    .addComponent(txtDirectorio, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        btnListaAuto.setFont(new java.awt.Font("Helvetica", 0, 13)); // NOI18N
        btnListaAuto.setText("Explorar Archivos");
        btnListaAuto.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnListaAutoActionPerformed(evt);
            }
        });

        txtListaAuto.setFont(new java.awt.Font("Helvetica Neue", 2, 12)); // NOI18N
        txtListaAuto.setText("Selecciona el archivo PDF con la lista  autógrafa de asistencia...");
        txtListaAuto.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txtListaAuto.setEnabled(false);
        txtListaAuto.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                txtListaAutoFocusLost(evt);
            }
        });
        txtListaAuto.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                txtListaAutoMouseClicked(evt);
            }
        });
        txtListaAuto.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtListaAutoActionPerformed(evt);
            }
        });

        lblListaAuto.setFont(new java.awt.Font("Helvetica Neue", 1, 14)); // NOI18N
        lblListaAuto.setText("¿Donde se encuentra la lista autógrafa de participantes?");

        lblLista.setFont(new java.awt.Font("Helvetica Neue", 1, 14)); // NOI18N
        lblLista.setText("¿Donde se encuentra la lista digital de participantes?");

        btnLista.setFont(new java.awt.Font("Helvetica", 0, 13)); // NOI18N
        btnLista.setText("Explorar Archivos");
        btnLista.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnListaActionPerformed(evt);
            }
        });

        txtLista.setFont(new java.awt.Font("Helvetica Neue", 2, 12)); // NOI18N
        txtLista.setText("Selecciona el archivo Excel con la lista de asistencia...");
        txtLista.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txtLista.setEnabled(false);
        txtLista.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                txtListaFocusLost(evt);
            }
        });
        txtLista.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                txtListaMouseClicked(evt);
            }
        });
        txtLista.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtListaActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGap(16, 16, 16)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblLista)
                            .addComponent(lblListaAuto)))
                    .addComponent(txtLista, javax.swing.GroupLayout.PREFERRED_SIZE, 476, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtListaAuto, javax.swing.GroupLayout.DEFAULT_SIZE, 478, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnLista)
                    .addComponent(btnListaAuto))
                .addContainerGap())
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lblLista)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtLista, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnLista))
                .addGap(27, 27, 27)
                .addComponent(lblListaAuto)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtListaAuto, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnListaAuto))
                .addGap(6, 6, 6))
        );

        jPanel4.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        btnExplorarCarpetas.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        btnExplorarCarpetas.setText("Explorar Carpetas");
        btnExplorarCarpetas.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExplorarCarpetasActionPerformed(evt);
            }
        });

        txtGuardarDocs.setFont(new java.awt.Font("Helvetica Neue", 2, 12)); // NOI18N
        txtGuardarDocs.setText("Selecciona carpeta destino...");
        txtGuardarDocs.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txtGuardarDocs.setEnabled(false);

        lblDirectorioGuardar.setFont(new java.awt.Font("Helvetica Neue", 1, 14)); // NOI18N
        lblDirectorioGuardar.setText("¿Donde se van a guardar los documentos generados?");

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addComponent(lblDirectorioGuardar)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addComponent(txtGuardarDocs)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(btnExplorarCarpetas, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(lblDirectorioGuardar)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtGuardarDocs, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnExplorarCarpetas))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel5.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        txtRegistro.setFont(new java.awt.Font("Helvetica Neue", 2, 12)); // NOI18N
        txtRegistro.setHorizontalAlignment(javax.swing.JTextField.LEFT);
        txtRegistro.setText("Si deseas cambiar el archivo PDF de registro default seleccionalo aquí");
        txtRegistro.setToolTipText("");
        txtRegistro.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txtRegistro.setDoubleBuffered(true);
        txtRegistro.setEnabled(false);
        txtRegistro.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtRegistroFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                txtRegistroFocusLost(evt);
            }
        });
        txtRegistro.addInputMethodListener(new java.awt.event.InputMethodListener() {
            public void inputMethodTextChanged(java.awt.event.InputMethodEvent evt) {
                txtRegistroInputMethodTextChanged(evt);
            }
            public void caretPositionChanged(java.awt.event.InputMethodEvent evt) {
            }
        });
        txtRegistro.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtRegistroActionPerformed(evt);
            }
        });
        txtRegistro.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                txtRegistroPropertyChange(evt);
            }
        });

        btnExplorarRegistro.setFont(new java.awt.Font("Helvetica Neue", 0, 13)); // NOI18N
        btnExplorarRegistro.setText("Explorar Archivos");
        btnExplorarRegistro.setEnabled(false);
        btnExplorarRegistro.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExplorarRegistroActionPerformed(evt);
            }
        });

        chkRegistro.setFont(new java.awt.Font("Helvetica", 3, 14)); // NOI18N
        chkRegistro.setText("¿Deseas cambiar el archivo registro de Jorge Razon 2016?");
        chkRegistro.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chkRegistroActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(chkRegistro, javax.swing.GroupLayout.PREFERRED_SIZE, 533, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(txtRegistro)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(btnExplorarRegistro)))
                .addContainerGap())
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addComponent(chkRegistro)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnExplorarRegistro)
                    .addComponent(txtRegistro, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(generar)
                        .addGap(18, 18, 18)
                        .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(19, 19, 19))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(13, 13, 13)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(generar, javax.swing.GroupLayout.DEFAULT_SIZE, 84, Short.MAX_VALUE)
                            .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
                .addGap(0, 19, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    private void logAccess() throws MalformedURLException, ProtocolException, IOException{
        
        String hostname = "Unknown";
        String username = "capacitacion";
        Format formatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
        String formatedDate = formatter.format(new Date());
                
        try
        {
            InetAddress addr;
            addr = InetAddress.getLocalHost();
            hostname = addr.getHostName();
        }
        catch (UnknownHostException ex)
        {
            System.out.println("Hostname can not be resolved");
        }
        
        String url = "https://script.google.com/macros/s/AKfycbxkZelT2ayE7qQ4r4SdOFOJ9GBNkKxncHReMrwNuUfQfHJ-PcA/exec?PCname="+hostname+"&username="+username+"&d="+formatedDate;
        
        URL obj = new URL(url);
        HttpURLConnection con = (HttpURLConnection) obj.openConnection();

        // optional default is GET
        con.setRequestMethod("GET");

        //add request header
        con.setRequestProperty("User-Agent", "Mozilla/5.0");
        int responseCode = 0;
        
        try{
            responseCode = con.getResponseCode();
        }catch(Exception ex){
            System.out.println(ex.getMessage());
        }
        
        System.out.println("response code: " + responseCode);

        if (responseCode != 200){

            String s = hostname + "%capacitacion%" + formatedDate+"\n";
            
            Files.write(Paths.get("./log.txt"), s.getBytes(), StandardOpenOption.APPEND);
 
        } else {
            File file = new File("./log.txt");
            
            BufferedReader br = new BufferedReader(new FileReader(file));     
            
            String line;
            String [] data;
            ArrayList <String> newData = new ArrayList<>();
            ListIterator <String> it = newData.listIterator();
            
            while ((line = br.readLine()) != null) {
               data = line.split("%");
               
               if (data.length == 3){
                   hostname = data[0];
                   username = data[1];
                   formatedDate = data[2];
               }
                url = "https://script.google.com/macros/s/AKfycbxkZelT2ayE7qQ4r4SdOFOJ9GBNkKxncHReMrwNuUfQfHJ-PcA/exec?PCname="+hostname+"&username="+username+"&d="+formatedDate;
        
                obj = new URL(url);
                con = (HttpURLConnection) obj.openConnection();

                // optional default is GET
                con.setRequestMethod("GET");

                //add request header
                con.setRequestProperty("User-Agent", "Mozilla/5.0");

                if (con.getResponseCode() != 200){
                    break;
                }else{
                    newData.add(line);                    
                }                
            }
            
            FileWriter fw = new FileWriter(file.getAbsoluteFile());           
            fw.write("");
            while(it.hasNext()){
                String nl = it.next();
                fw.append(nl);            
            }
  
        }
    }
    
    private void generarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_generarActionPerformed
        
       
        String directorioPath, destFolder, listaPath;
        String chkDip,chkDipFirma,chkDipLogo,chkConst,chkConstFirma,chkConstLogo,chkDC3a,chkDC3Firmaa,chkDC3Logoa,chkCompilado;
        String unidadCapacitadora,instructor,registro,lista_autografa;
        
        directorioPath = this.txtDirectorio.getText();
        if (directorioPath.startsWith("Selecciona")){
            directorioPath = "";
        }
        destFolder = this.txtGuardarDocs.getText();
        if (destFolder.startsWith("Selecciona")){
            destFolder = "";
        }
        listaPath = this.txtLista.getText();
        if (listaPath.startsWith("Selecciona")){
            listaPath = "";
        }
        unidadCapacitadora = (String)this.cboxUnidadCapacitadora.getSelectedItem();
        instructor = (String)this.cboxInstructor.getSelectedItem();
        
        chkDip = String.valueOf(this.chkDiplomas.isSelected());
        chkDipFirma = String.valueOf(this.chkDiplomaFirma.isSelected());
        chkDipLogo = String.valueOf(this.chkDiplomasLogo.isSelected());
        
        chkConst = String.valueOf(this.chkConstancias.isSelected());
        chkConstFirma = String.valueOf(this.chkConstanciasFirma1.isSelected());
        chkConstLogo = String.valueOf(this.chkConstanciasLogo.isSelected());
        
        chkDC3a = String.valueOf(this.chkDC3.isSelected());
        chkDC3Firmaa = String.valueOf(this.chkDC3Firma.isSelected());
        chkDC3Logoa = String.valueOf(this.chkDC3Logo.isSelected());
        
        chkCompilado = String.valueOf(this.chkArchivoCompilado.isSelected());
        
        registro = this.txtRegistro.getText();
         if (registro.startsWith("Si deseas")){
            registro = "";
        }
         
        lista_autografa = this.txtListaAuto.getText();
        if (lista_autografa.startsWith("Selecciona")){
            lista_autografa = "";
        }
        
        String[] args = {directorioPath, destFolder, listaPath,unidadCapacitadora,instructor,chkDip,chkDipFirma,chkDipLogo,chkConst,chkConstFirma,chkConstLogo,chkDC3a,chkDC3Firmaa,chkDC3Logoa,chkCompilado,registro,lista_autografa};

        
            SwingWorker sw = new SwingWorker(){
                @Override
                protected void done() {
                    // Close the dialog
                    generar.setText("Generar Documentos");
                    generar.setEnabled(true);                                                   
                }

                @Override
                protected Object doInBackground() throws Exception {
                    // Do the long running task here
                    // Call "publish()" to pass the data to "process()"
                    // return something meaningful                   
                    cdiscisa.Cdiscisa.main(args);
                    return null;
                }
            };
            this.generar.setText("Generando Documentos...");
            this.generar.setVisible(true);
            this.generar.setEnabled(false);            
            SwingUtilities.updateComponentTreeUI(this);
            this.generar.repaint();
            
            sw.execute();

    }//GEN-LAST:event_generarActionPerformed

    private void btnExplorarDirectorioActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExplorarDirectorioActionPerformed
        // TODO add your handling code here:
        
        
        JFileChooser openFile = new JFileChooser();
        openFile.setApproveButtonText("Seleccionar");
        openFile.setCurrentDirectory(new File(System.getProperty("user.home")));
        FileNameExtensionFilter fEx = new FileNameExtensionFilter("Excel (xlsx)","xlsx");
        openFile.setFileFilter(fEx);
        
        if (openFile.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
        {
            txtDirectorio.setText(openFile.getSelectedFile().getAbsolutePath());
        }
        
            
    }//GEN-LAST:event_btnExplorarDirectorioActionPerformed

    private void btnExplorarCarpetasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExplorarCarpetasActionPerformed
        // TODO add your handling code here:
        JFileChooser openFile2 = new JFileChooser();
        openFile2.setApproveButtonText("Seleccionar");
        openFile2.setCurrentDirectory(new File(System.getProperty("user.home")));
        openFile2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        if (openFile2.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
        {
        txtGuardarDocs.setText(openFile2.getSelectedFile().getAbsolutePath());
        }
    }//GEN-LAST:event_btnExplorarCarpetasActionPerformed

    private void btnListaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnListaActionPerformed
        // TODO add your handling code here:
        JFileChooser openFile3 = new JFileChooser();
        openFile3.setApproveButtonText("Seleccionar");
        openFile3.setCurrentDirectory(new File(System.getProperty("user.home")));
        
        FileNameExtensionFilter fEx = new FileNameExtensionFilter("Excel","xlsx");
        openFile3.setFileFilter(fEx);
        
        if (openFile3.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
        {
        txtLista.setText(openFile3.getSelectedFile().getAbsolutePath());
        }
    }//GEN-LAST:event_btnListaActionPerformed

    private void chkDiplomasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chkDiplomasActionPerformed
        // TODO add your handling code here:
        /*
        if (this.chkDiplomas.isSelected()){
            this.chkDiplomasLogo.setSelected(true);
            this.chkDiplomasLogo.setEnabled(true);
            this.chkDiplomaFirma.setSelected(true);
            this.chkDiplomaFirma.setEnabled(true);
        } else {
            this.chkDiplomasLogo.setSelected(false);
            this.chkDiplomasLogo.setEnabled(false);
            this.chkDiplomaFirma.setSelected(false);
            this.chkDiplomaFirma.setEnabled(false);
        }*/
    }//GEN-LAST:event_chkDiplomasActionPerformed

    private void chkConstanciasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chkConstanciasActionPerformed
        // TODO add your handling code here:
        /*if (this.chkConstancias.isSelected()){
            this.chkConstanciasLogo.setSelected(true);
            this.chkConstanciasLogo.setEnabled(true);
            this.chkConstanciasFirma1.setSelected(true);
            this.chkConstanciasFirma1.setEnabled(true);
        } else {
            this.chkConstanciasLogo.setSelected(false);
            this.chkConstanciasLogo.setEnabled(false);
            this.chkConstanciasFirma1.setSelected(false);
            this.chkConstanciasFirma1.setEnabled(false);
        }*/
    }//GEN-LAST:event_chkConstanciasActionPerformed

    private void chkDC3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chkDC3ActionPerformed
        // TODO add your handling code here:
        /*
        if (this.chkDC3.isSelected()){
            this.chkDC3Firma.setSelected(true);
            this.chkDC3Firma.setEnabled(true);
            this.chkDC3Logo.setSelected(true);
            this.chkDC3Logo.setEnabled(true);
        } else {
            this.chkDC3Firma.setSelected(false);
            this.chkDC3Firma.setEnabled(false);
            this.chkDC3Logo.setSelected(false);
            this.chkDC3Logo.setEnabled(false);
        } */
        
    }//GEN-LAST:event_chkDC3ActionPerformed

    private void cboxUnidadCapacitadoraActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cboxUnidadCapacitadoraActionPerformed
        // TODO add your handling code here:
        
        if (cboxUnidadCapacitadora.getSelectedIndex() == 0){
            cboxInstructor.removeAllItems();
            cboxInstructor.addItem("TSI. Jorge Antonio Razón Gil");
        } else {
            cboxInstructor.removeAllItems();
            cboxInstructor.addItem("TSI. Jorge Antonio Razón Gil");
            cboxInstructor.addItem("Ing. Jorge Antonio Razón Gutierrez");
            cboxInstructor.addItem("Manuel Anguiano Razón");
        }
    }//GEN-LAST:event_cboxUnidadCapacitadoraActionPerformed

    private void txtDirectorioInputMethodTextChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_txtDirectorioInputMethodTextChanged
        // TODO add your handling code here:
       
            
    }//GEN-LAST:event_txtDirectorioInputMethodTextChanged

    private void txtDirectorioPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_txtDirectorioPropertyChange
        // TODO add your handling code here:
        
         
    }//GEN-LAST:event_txtDirectorioPropertyChange

    private void txtDirectorioFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtDirectorioFocusLost
        // TODO add your handling code here:
        /*File f;
        String path;
        path = this.txtDirectorio.getText();
        
        f = new File(path);
            if(!f.exists() || f.isDirectory()) { 
                JOptionPane.showMessageDialog(null, "El archivo de no se encuentra en la ruta seleccionada");
                this.txtDirectorio.setText("Selecciona el archivo que continene las direcciones de las sucursales");
            } 
        */
        
    }//GEN-LAST:event_txtDirectorioFocusLost

    private void txtDirectorioFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtDirectorioFocusGained
        // TODO add your handling code here:
        //this.txtDirectorio.setText("");
    }//GEN-LAST:event_txtDirectorioFocusGained

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        // TODO add your handling code here:
        //this.txtDirectorio.setText(System.getProperty("user.dir"));
        this.txtGuardarDocs.setText(System.getProperty("user.dir"));
        //this.txtLista.setText(System.getProperty("user.dir"));
        
    }//GEN-LAST:event_formWindowOpened

    private void txtDirectorioActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtDirectorioActionPerformed
        // TODO add your handling code here:
        this.txtDirectorio.setText("");
    }//GEN-LAST:event_txtDirectorioActionPerformed

    private void txtListaFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtListaFocusLost
        // TODO add your handling code here:
        /*File f;
        String path;
        path = this.txtLista.getText();
        
        f = new File(path);
            if(!f.exists() || f.isDirectory()) { 
                JOptionPane.showMessageDialog(null, "El archivo de no se encuentra en la ruta seleccionada");
                this.txtLista.setText("Selecciona el archivo que continene la lista de asistencia");
            } 
        */
    }//GEN-LAST:event_txtListaFocusLost

    private void txtListaMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_txtListaMouseClicked
        // TODO add your handling code here:
        
    }//GEN-LAST:event_txtListaMouseClicked

    private void txtListaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtListaActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtListaActionPerformed

    private void chkArchivoCompiladoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chkArchivoCompiladoActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_chkArchivoCompiladoActionPerformed

    private void txtListaAutoFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtListaAutoFocusLost
        // TODO add your handling code here:
    }//GEN-LAST:event_txtListaAutoFocusLost

    private void txtListaAutoMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_txtListaAutoMouseClicked
        // TODO add your handling code here:
    }//GEN-LAST:event_txtListaAutoMouseClicked

    private void txtListaAutoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtListaAutoActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtListaAutoActionPerformed

    private void btnListaAutoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnListaAutoActionPerformed
        // TODO add your handling code here:
        JFileChooser openFile4 = new JFileChooser();
        openFile4.setApproveButtonText("Seleccionar");
        openFile4.setCurrentDirectory(new File(System.getProperty("user.home")));
        
        FileNameExtensionFilter fEx = new FileNameExtensionFilter("PDF","pdf");
        openFile4.setFileFilter(fEx);
        
        if (openFile4.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
        {
        txtListaAuto.setText(openFile4.getSelectedFile().getAbsolutePath());
        }
        
    }//GEN-LAST:event_btnListaAutoActionPerformed

    private void txtRegistroFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtRegistroFocusGained
        // TODO add your handling code here:
    }//GEN-LAST:event_txtRegistroFocusGained

    private void txtRegistroFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtRegistroFocusLost
        // TODO add your handling code here:
    }//GEN-LAST:event_txtRegistroFocusLost

    private void txtRegistroInputMethodTextChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_txtRegistroInputMethodTextChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_txtRegistroInputMethodTextChanged

    private void txtRegistroActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtRegistroActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtRegistroActionPerformed

    private void txtRegistroPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_txtRegistroPropertyChange
        // TODO add your handling code here:
    }//GEN-LAST:event_txtRegistroPropertyChange

    private void btnExplorarRegistroActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExplorarRegistroActionPerformed
        // TODO add your handling code here:
        JFileChooser openFile5 = new JFileChooser();
        openFile5.setApproveButtonText("Seleccionar");
        openFile5.setCurrentDirectory(new File(System.getProperty("user.home")));
        
        FileNameExtensionFilter fEx = new FileNameExtensionFilter("PDF","pdf");
        openFile5.setFileFilter(fEx);
        
        if (openFile5.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
        {
        txtRegistro.setText(openFile5.getSelectedFile().getAbsolutePath());
        }
        
    }//GEN-LAST:event_btnExplorarRegistroActionPerformed

    private void chkRegistroActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chkRegistroActionPerformed
        // TODO add your handling code here:
       
        if (this.chkRegistro.isSelected()){                                    
            this.btnExplorarRegistro.setEnabled(true);
        } else {                                   
            this.btnExplorarRegistro.setEnabled(false);
        }
        
    }//GEN-LAST:event_chkRegistroActionPerformed

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
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    new MainForm().setVisible(true);
                } catch (IOException ex) {
                    Logger.getLogger(MainForm.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnExplorarCarpetas;
    private javax.swing.JButton btnExplorarDirectorio;
    private javax.swing.JButton btnExplorarRegistro;
    private javax.swing.JButton btnLista;
    private javax.swing.JButton btnListaAuto;
    private javax.swing.JComboBox<String> cboxInstructor;
    private javax.swing.JComboBox<String> cboxUnidadCapacitadora;
    private javax.swing.JCheckBox chkArchivoCompilado;
    private javax.swing.JCheckBox chkConstancias;
    private javax.swing.JCheckBox chkConstanciasFirma1;
    private javax.swing.JCheckBox chkConstanciasLogo;
    private javax.swing.JCheckBox chkDC3;
    private javax.swing.JCheckBox chkDC3Firma;
    private javax.swing.JCheckBox chkDC3Logo;
    private javax.swing.JCheckBox chkDiplomaFirma;
    private javax.swing.JCheckBox chkDiplomas;
    private javax.swing.JCheckBox chkDiplomasLogo;
    private javax.swing.JCheckBox chkRegistro;
    private javax.swing.JButton generar;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JLabel lblCheckDocumentos;
    private javax.swing.JLabel lblDirectorio;
    private javax.swing.JLabel lblDirectorioGuardar;
    private javax.swing.JLabel lblInstructor;
    private javax.swing.JLabel lblLista;
    private javax.swing.JLabel lblListaAuto;
    private javax.swing.JLabel lblUnidadCapacitadora;
    private javax.swing.JTextField txtDirectorio;
    private javax.swing.JTextField txtGuardarDocs;
    private javax.swing.JTextField txtLista;
    private javax.swing.JTextField txtListaAuto;
    private javax.swing.JTextField txtRegistro;
    // End of variables declaration//GEN-END:variables
}
