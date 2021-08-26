package Interface;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.table.DefaultTableModel;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Principal extends javax.swing.JFrame {

    /**
     * Documentação: Canal do Youtube:
     * https://www.youtube.com/watch?v=-Gc3_ViFNb4
     */
//criando um objeto File que irá fazer a leitura do objeto no nosso caso a planilha do excel 
    File file;
    
    Font fonteTitulo = new Font("Tahoma", Font.BOLD, 16);
    Font fontePadrao = new Font("Tahoma", Font.BOLD, 12);
    

//criar um objeto do tipo workbook que permite manipular objetos 
    Workbook workbook;

    public Principal() {
        initComponents();
        //chamar a função sem retorno void frontEndSwing
        frontEndSwing();
    }

    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        PainelPrincipal = new javax.swing.JPanel();
        jtlocal = new javax.swing.JTextField();
        btnXLS = new javax.swing.JButton();
        lblLogoTrybe = new javax.swing.JLabel();
        ScPainelTablePrincipal = new javax.swing.JScrollPane();
        tabela = new javax.swing.JTable();
        jTitulo = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        PainelPrincipal.setBackground(java.awt.Color.white);
        PainelPrincipal.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        PainelPrincipal.add(jtlocal, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 110, 400, 40));

        btnXLS.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnXLSActionPerformed(evt);
            }
        });
        PainelPrincipal.add(btnXLS, new org.netbeans.lib.awtextra.AbsoluteConstraints(430, 110, 45, 45));
        PainelPrincipal.add(lblLogoTrybe, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 10, 160, 90));

        tabela.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Item", "Descrição"
            }
        ));
        ScPainelTablePrincipal.setViewportView(tabela);
        if (tabela.getColumnModel().getColumnCount() > 0) {
            tabela.getColumnModel().getColumn(0).setMinWidth(50);
            tabela.getColumnModel().getColumn(0).setMaxWidth(50);
        }

        PainelPrincipal.add(ScPainelTablePrincipal, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 190, 450, 360));

        jTitulo.setFont(new java.awt.Font("Ubuntu", 1, 18)); // NOI18N
        jTitulo.setForeground(javax.swing.UIManager.getDefaults().getColor("ComboBox.disabledForeground"));
        PainelPrincipal.add(jTitulo, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 160, 450, 30));

        getContentPane().add(PainelPrincipal, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 500, 580));

        setSize(new java.awt.Dimension(511, 613));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    
    private void frontEndSwing() {

        lblLogoTrybe.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/trybePadrão.png")));
        btnXLS.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/excel.png")));
        

    }
    
    
    
    
    
    
    
    private void btnXLSActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnXLSActionPerformed

        //// inicio do codigo
        JFileChooser fc = new JFileChooser();
        //por meio do objeto [fc] criando iremos agora dimensionar o tamanho da nossa tela 
        fc.setPreferredSize(new Dimension(750, 400));
        //com o código abaixo iremos selecionar um único arquivo por vez do tipo xls
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
        //aqui criamos um objeto int que recebe o resultado do clique na janela aberta 
        int resultado = fc.showOpenDialog(this);

        //se JFileChooser receber cancel da janela de dialogo ele irá fechar a janela e não irá trazer 
        //resultado nenhum 
        if (resultado == JFileChooser.CANCEL_OPTION) {
            //para cancelar a janela se necessário caso contrário se for setado algum arquivo ele irá 
            //trazer o caminho do arquivo setado 
        } else {
            //pega o caminho do arquivo setado 
            file = fc.getSelectedFile();
            //tira os espaços em brnaco 
            jtlocal.setText(file.toString().trim());
            // método preenche titulo e tabela 
            preencherJtableETitulo();

        }

    }//GEN-LAST:event_btnXLSActionPerformed

    
    private void preencherJtableETitulo(){
    
    
        try {

            /**
             * criando um objeto do tipo workbook, que irá pegar o caminho
             * capturado pelo objeto JFileChooser (fc) objeto esse que está
             * vinculado ao botão XLS disposto neste JForm
             */
            workbook = Workbook.getWorkbook(new File(jtlocal.getText().trim()));

        } catch (IOException | BiffException ex) {
         
            System.out.println("Erro:"+ex.getMessage());
            ex.printStackTrace();
        }

        /**
         * Agora iremos criar um objeto do tipo sheet que irá pegar cada umas
         * das abas da planinha do excel, geralmente por padrão são abertas
         * 3(tres...Plan1, Plan2, Plan3). em seguida iremos setar neste objeto
         * sheet a planilha que iremos de fato trabalhar neste caso em
         * especifico a primeira planilha Plan1, pois, o índice setado no metodo
         * getSheet foi o 0(zero) neste caso a primeira posição logo a Plan1 do
         * Excel.
         */
        Sheet sheet = workbook.getSheet(0);

        /**
         * Criando um objeto do tipo Cell que irá capturar a primeira Célula da
         * planilha acima setada em nosso caso Plan1(Excel). Aqui funciona o
         * princícipio de Matriz quando comparada com a Planilha do Excel Linha
         * e Coluna no caso abaixo especificado getContents(0,0) Coluna 0(zero)
         * e Linha 0(zero)
         */
        Cell c0 = sheet.getCell(0, 0);

        /**
         * Agora pegamos o objeto contendo as informações capturadas na (Plan1)
         * e c0(0,0) e setamos no objeto do tipo String titulo a informação
         * capturada
         */
        String titulo = c0.getContents();

        /**
         * em seguida setamos essa informação contida agora em título no
         * jText(jttitulo) por meio do método setText();
         */
        jTitulo.setText(titulo);
        jTitulo.setFont(fonteTitulo);
        jTitulo.setForeground(Color.GREEN);
        /**
         * Nesse ponto irei contar o número de linhas da minha planilha para só
         * então fazer a interação da mesma for meio de um laço de repetição for
         */
        int linhas = sheet.getRows();

        ////inicio do for
        for (int i = 2; i < linhas; i++) {

            /**
             * Dentro do laço de repetição a primeira informaçãoa ser pega é a
             * contida na posição 0(zero) e linha 2, pois, for do laço tem sua
             * interação inicializada na posição 2 como descrita logo acima
             * (essas posições pegas aqui estão descrita do ponto de vista da
             * programação indices zero primeira posição da planilha)
             *
             */
            Cell ca = sheet.getCell(0, i);

            /**
             * a segunda informação a ser pela é na posição 1(um) linha 2 (isso
             * do ponto de vista da programação trabalhando com indices)
             */
            Cell cb = sheet.getCell(1, i);

            /**
             * Criamos duas variáveis e colocamos o conteúdo capturado de
             * ca(Coluna a ) em item e cb(Coluna b) em desc
             */
            String item = ca.getContents();
            String desc = cb.getContents();

            /**
             * Agora instaciamos um objeto mp do tipo DefaultTableModel e
             * colocamos numa tabela chamada tb1
             */
            DefaultTableModel mp = (DefaultTableModel) tabela.getModel();

            /**
             * Vou adicionar uma linha que conterá informações dos objetos item
             * desc capturados acima
             */
            mp.addRow(new String[]{item, desc});

        }   //////// fim do for

        workbook.close();

//////fim do codigo
    
    
    }
    
    
    
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
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Principal().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JPanel PainelPrincipal;
    private javax.swing.JScrollPane ScPainelTablePrincipal;
    private javax.swing.JButton btnXLS;
    private javax.swing.JLabel jTitulo;
    private javax.swing.JTextField jtlocal;
    private javax.swing.JLabel lblLogoTrybe;
    private javax.swing.JTable tabela;
    // End of variables declaration//GEN-END:variables
}
