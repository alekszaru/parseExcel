import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.*;
import java.util.ArrayList;

public class Main {
    private static String dtsFileName = "src/test/DTS.xls";
    private static String eaFileName = "src/test/EA.xls";
    private static String eaPriceFileName = "src/test/�� prise.xls";
    private static String mkPriceFileName = "";
    private static String mkFileName = "src/test/MK.xls";
    private static String bkzPriceFileName = "";
    private static String kpkzPriceFileName = "src/test/KPKZ.xls";


    public static ArrayList<String> rezult = new ArrayList<String>();
    private static double eaAluminiumDiscount = 0.92;
    private static double eaCuprumDiscount = 0.96;


    public static void main(String[] args) throws IOException {

        SimpleGUI app = new SimpleGUI();
        app.setVisible(true);

    }


    public static void findMatchesDTS(String file, String request) throws IOException { //����� � ����� ���
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row;

//        System.out.println("----------------------------------------------------------");
//        System.out.println(" � � � ");

        if (request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())) {
                        Double cuantaty = row.getCell(3).getNumericCellValue();
                        Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " �  " + String.format("%.2f", price / 1000) + "���/�";
                        //System.out.println(answer);
                        rezult.add(answer + " ���");
                    }
                } catch (IllegalStateException e) {
                } catch (NullPointerException e5) {
                }
            }
        } else {
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase())) {
                        Double cuantaty = row.getCell(3).getNumericCellValue();
                        Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " �  " + String.format("%.2f", price / 1000) + "���/�";
                        // System.out.println(answer);
                        rezult.add(answer + " ���");
                    }
                } catch (IllegalStateException e) {
                } catch (NullPointerException e5) {
                }
            }
        }


        myExcelBook.close();


    } //����� � ����� ���


    public static void findMatchesEA(String file, String request) throws IOException {// ����� � ����� ������������
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row;

//        System.out.println("----------------------------------------------------------");
//        System.out.println(" ������������ ");

        if (request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);

            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase()) && (row.getCell(2).getNumericCellValue()) > 0) {

                        Double cuantaty = row.getCell(7).getNumericCellValue();

                        String answer = name + " - " + String.format("%.0f", cuantaty) + " � ";
                        //System.out.println(answer);
                        rezult.add(answer + " ������������");
                    }
                } catch (IllegalStateException e) {
                } catch (NullPointerException e2) {
                }
            }
        } else {
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase()) && (row.getCell(2).getNumericCellValue()) > 0) {
                        Double cuantaty = row.getCell(7).getNumericCellValue();
                        String answer = name + " - " + String.format("%.0f", cuantaty) + " � ";

                        //System.out.println(answer);
                        rezult.add(answer + " ������������");
                    }
                } catch (IllegalStateException e) {
                } catch (NullPointerException e2) {
                }
            }
        }
        myExcelBook.close();
    } // ����� � ����� ������������


    public static void findMatchesKPKZ(String file, String request) throws IOException {//����� � ������ ��������� �����
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row;

        for (int i = 0; i < myExcelSheet.getPhysicalNumberOfRows() ; i++) {
            row = myExcelSheet.getRow(i);
            String fullName = "";
            String temp = "";

            try {

                fullName += row.getCell(0).getStringCellValue() + " ";

                if((temp = row.getCell(3).getStringCellValue())!="")
                fullName += temp + "x";

                if((temp = row.getCell(5).getStringCellValue())!="")
                    fullName += temp + "x";

                if((temp = row.getCell(10).getStringCellValue())!="")
                    fullName += "+" + temp + "x";

                if((temp = row.getCell(12).getStringCellValue())!="")
                    fullName += temp;
// STOP_POINT


if(fullName.contains(request))
                System.out.println(fullName);

//                if (type1.toUpperCase().contains(type.toUpperCase()) && c11.toUpperCase().contains(c1.toUpperCase()) && c21.toUpperCase().contains(c2.toUpperCase())) {
//                    Double price = row.getCell(15).getNumericCellValue() / 1000;
//
//                    String answer = type1 + " " + c11 + "x" + c21 + "   " + String.format("%.2f", price) + " ���/�";
//
//
//                    System.out.println(answer);
//                    rezult.add(answer + " ����� ��������� �����");
//                }
            } catch (IllegalStateException e) {
            } catch (NullPointerException e1) {
            }
        }
//
//  //         if (request.contains(" ")&& (request.toUpperCase().contains("�"))) {
////                String request1 = request.substring(0, request.lastIndexOf(" "));
////                System.out.println("request1 " +request1);
////                String request2 = request.substring(request.toUpperCase().lastIndexOf(" ")+1);
////                System.out.println("request2 " +request2);
////                String request3 = request.substring(request.toUpperCase().indexOf("X")+1);
////                System.out.println("request3 "+ request3);
////
////                 for (int i = 0; i < myExcelSheet.getPhysicalNumberOfRows() ; i++) {
////                        row = myExcelSheet.getRow(i);
////                        try {
////                            String name = row.getCell(0 ).getStringCellValue();
////
////                            String name2 = row.getCell(3 ).getStringCellValue();
////
////                            String name3 = row.getCell(5).getStringCellValue();
////
////                            if (name.toUpperCase().contains(request1.toUpperCase()) && request2.toUpperCase().contains(name2.toUpperCase())&& request2.toUpperCase().contains(name2.toUpperCase())) {
////                                Double price = row.getCell(15 ).getNumericCellValue()/1000;
////
////                                   String answer = name + " " + name2+ "x"+name3+ "   " + String.format("%.2f", price) + " ���/�";
////
////
////                                System.out.println(answer);
////                                rezult.add(answer+ " ����� ��������� �����");
////                            }
////                        } catch (IllegalStateException e) {  }
////                        catch (NullPointerException e1){  }
////                 }
////
// //   }
//
////            else{
////
////                  for (int i = 11; i < myExcelSheet.getPhysicalNumberOfRows() ; i++) {
////                        row = myExcelSheet.getRow(i);
////                        try {
////                            String name =row.getCell(0 ).getStringCellValue();
////                            String name2 =row.getCell(3 ).getStringCellValue();
////                            String name3 =row.getCell(5 ).getStringCellValue();
////
////                            if (name.toUpperCase().contains(request.toUpperCase()) ) {
////                                Double price = row.getCell(15 ).getNumericCellValue()/1000;
////
////                                String answer = name + " " + name2+ "x"+name3+ "   " + String.format("%.2f", price) + " ���/�";
////
////                                //System.out.println(answer);
////                                rezult.add(answer + "����� ��������� �����");
////                            }
////                        } catch (IllegalStateException e) {
////                        }
////                        catch (NullPointerException e1){
////                        }
////                  }
////
////            }
//
} //����� � ������ ��������� �����


    public static void findMatchesMK(String file, String request) throws IOException{ //����� � ����� ������ ������

        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row ;

//        System.out.println("----------------------------------------------------------");
//        System.out.println(" ������ ������");

        if(request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())) {
                        Double cuantaty = row.getCell(5).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000)+ " � ";
                        //System.out.println(answer);
                        rezult.add(answer+ " ������ ������");
                    }
                } catch (IllegalStateException e) {}
                catch (NullPointerException e5){}
            }
        }

        else{
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase())) {
                        Double cuantaty = row.getCell(5).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000)+ " � ";
                        //System.out.println(answer);
                        rezult.add(answer + " ������ ������");
                    }
                }
                catch (IllegalStateException e){}
                catch (NullPointerException e5){}
            }
        }
        myExcelBook.close();
    } //����� � ����� ������ ������


    public static void findMatchesInPriceEA(String file, String request) throws IOException {//����� � ������ ������������
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        int numsOfSheets = myExcelBook.getNumberOfSheets();
        if(request.indexOf(" ")==0){request=request.substring(1);}
//        System.out.println("----------------------------------------------------------");
//        System.out.println("����� ������������" );

        for(int s=0; s<numsOfSheets; s++){

            HSSFSheet myExcelSheet = myExcelBook.getSheetAt(s);
            HSSFRow row;
            if (request.contains(" ")) {
                String request1 = request.substring(0, request.lastIndexOf(" "));
                String request2 = request.substring(request.lastIndexOf(" ") + 1);

                for (int j = 0; j < 5; j++) {
                    for (int i = 1; i < myExcelSheet.getPhysicalNumberOfRows() ; i++) {
                        row = myExcelSheet.getRow(i);
                        try {
                            String name = row.getCell(1 + 4 * j).getStringCellValue();
                            String name2 = row.getCell(2 + 4 * j).getStringCellValue();
                            if (name.toUpperCase().startsWith(request1.toUpperCase()) && name2.toUpperCase().contains(request2.toUpperCase())) {
                                Double price = row.getCell(3 + 4 * j).getNumericCellValue();
                                String answer ="";
                                if(name.toUpperCase().startsWith("�")||name.toUpperCase().startsWith("���")){
                                    answer = name + " " + name2+ "  " + String.format("%.2f", price*eaAluminiumDiscount ) + " ���/�";}
                                else {answer = name + " " + name2+ "  " + String.format("%.2f", price*eaCuprumDiscount ) + " ���/�";}

                                //System.out.println(answer);
                                rezult.add(answer+ " ����� ������������");
                            }
                        } catch (IllegalStateException e) {  }
                        catch (NullPointerException e1){  }
                    }
                }
            }

            else{

                for (int j = 0; j < 2; j++) {
                    for (int i = 1; i < myExcelSheet.getPhysicalNumberOfRows() ; i++) {
                        row = myExcelSheet.getRow(i);
                        try {
                            String name =row.getCell(1 + 4 * j).getStringCellValue();
                            String name2 =row.getCell(2 + 4 * j).getStringCellValue();
                            if (name.toUpperCase().startsWith(request.toUpperCase()) ) {
                                Double price = row.getCell(3 + 4 * j).getNumericCellValue();

                                String answer ="";
                                if(name.toUpperCase().startsWith("�")||name.toUpperCase().startsWith("���")){
                                    answer = name + " " + name2+ "  " + String.format("%.2f", price*eaAluminiumDiscount ) + " ���/�";}
                                else {answer = name + " " + name2+ "  " + String.format("%.2f", price*eaCuprumDiscount ) + " ���/�";}
                                //System.out.println(answer);
                                rezult.add(answer + "����� ������������");
                            }
                        } catch (IllegalStateException e) {
                        }
                        catch (NullPointerException e1){
                        }
                    }
                }
            }
        }
    } //����� � ������ ����������

    public static class SimpleGUI extends JFrame {
        private JButton button = new JButton("�����");
        public JTextField input = new JTextField("", 5);
        private JLabel label = new JLabel("������� ������ � ������� (���� 3�25)");
        public JTextArea textArea = new JTextArea("");
        public JButton button2 = new JButton("��������");



        public SimpleGUI() {
            super("����� ������� � �������� �����������");
            this.setBounds(500,200,400,500);
            this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

            JScrollPane jScrollPane = new JScrollPane(textArea);

            Container container = this.getContentPane();
            container.setLayout(null);


            label.setLocation(80,5);
            label.setSize(380,30);
            container.add(label);



            input.setLocation(10,35);
            input.setSize(370,30);
            input.setActionCommand("Enter");
            input.addActionListener(new ButtonEventListener());
            container.add(input);


            button.setLocation(10,75);
            button.setSize(370,30);
            button.setActionCommand("Button");
            button.addActionListener(new ButtonEventListener());
            container.add(button);


            jScrollPane.setLocation(10, 120);
            jScrollPane.setSize(370,300);
            container.add(jScrollPane);


            button2.setLocation(10,430);
            button2.setSize(370,30);
            button2.addActionListener(new ButtonEventListener());
            container.add(button2);

        }



        class ButtonEventListener implements ActionListener {

            public void actionPerformed(java.awt.event.ActionEvent evt) {
                String action = evt.getActionCommand();
                if(action.equals("Enter")||(action.equals("Button"))){
                    String str = input.getText();
                    try{
                        //findMatchesDTS(dtsFileName, str);
                       // findMatchesEA(eaFileName, str);
                      // findMatchesInPriceEA(eaPriceFileName, str);
                      //  findMatchesMK(mkFileName, str);
                        findMatchesKPKZ(kpkzPriceFileName,str);

                        }
                    catch (FileNotFoundException e){
                        System.out.println("��� �����");
                    }
                    catch (IOException e1){};

                    for(int i = 0; i<rezult.size();i++){
                        textArea.append(rezult.get(i)+"\n");
                    }
                    rezult.clear();
                }
                if(evt.getSource()==button2){
                    textArea.setText("");
                    rezult.clear();
                }
            }
        }



    }


}