

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;



public class Main extends Object {
    private static String dtsFileName = "D://stock/DTS.xls";
    private static String eaFileName = "D://stock/EA.xls";
    private static String eaPriceFileName = "D://stock/ЕА prise.xls";
    private static String mkPriceFileName = "";
    private static String mkFileName = "D://stock/MK.xls";
    private static String bkzPriceFileName = "";
    private static String request;

    public static ArrayList<String> getRezult() {
        return rezult;
    }

    public static ArrayList<String> rezult = new ArrayList<String>();
    private static double eaAluminiumDiscount = 0.92;
    private static double eaCuprumDiscount = 0.96;




    public static void main(String[] args) throws IOException {

//       BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
//      request = reader.readLine(



        SimpleGUI app = new SimpleGUI();
        app.setVisible(true);

        //System.out.println("==========================================================");




    }



    public static void findMatchesDTS(String file, String request) throws IOException{ //поиск в файле дтс
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row ;

        System.out.println("----------------------------------------------------------");
        System.out.println(" Д Т С ");

        if(request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())) {
                       Double cuantaty = row.getCell(3).getNumericCellValue();
                       Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " м  " + String.format("%.2f", price / 1000) + "грн/м";
                        System.out.println(answer);
                        rezult.add(answer+" ДТС");
                    }
                } catch (IllegalStateException e){}
                  catch (NullPointerException e5){}
            }
        }
        else{
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase())) {
                        Double cuantaty = row.getCell(3).getNumericCellValue();
                        Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " м  " + String.format("%.2f", price / 1000) + "грн/м";
                        System.out.println(answer);
                        rezult.add(answer+" ДТС");
                    }
                }
                    catch (IllegalStateException e){}
                    catch (NullPointerException e5){}
                }
        }


        myExcelBook.close();


    } //поиск в файле дтс


    public static void findMatchesEA(String file, String request) throws IOException{
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row ;

        System.out.println("----------------------------------------------------------");
        System.out.println(" ЭНЕРГОАЛЬЯНС ");

        if(request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);

            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())&&(row.getCell(2).getNumericCellValue())>0) {

                        Double cuantaty = row.getCell(7).getNumericCellValue();

                        String answer = name + " - " + String.format("%.0f", cuantaty) + " м ";
                        System.out.println(answer);
                        rezult.add(answer+" ЭНЕРГОАЛЬЯНС");
                    }
                } catch (IllegalStateException e) {}
                catch (NullPointerException e2){}
            }
        }

        else{
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase())&& (row.getCell(2).getNumericCellValue())>0) {
                        Double cuantaty = row.getCell(7).getNumericCellValue();
                        String answer = name + " - " + String.format("%.0f", cuantaty) + " м ";

                        System.out.println(answer);
                        rezult.add(answer+" ЭНЕРГОАЛЬЯНС");
                    }
                } catch (IllegalStateException e) {}
                catch (NullPointerException e2){}
            }
        }
        myExcelBook.close();
    } // поиск в файле Энергоальянс


    public static void findMatchesInPriceEA(String file, String request) throws IOException {//поиск в прайсе Энергоальянс
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        int numsOfSheets = myExcelBook.getNumberOfSheets();
        if(request.indexOf(" ")==0){request=request.substring(1);}
        System.out.println("----------------------------------------------------------");
        System.out.println("ПРАЙС ЭНЕРГОАЛЬЯНС" );

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
                            if (name.toUpperCase().contains(request1.toUpperCase()) && name2.toUpperCase().contains(request2.toUpperCase())) {
                                Double price = row.getCell(3 + 4 * j).getNumericCellValue();
                                String answer ="";
                                if(name.toUpperCase().startsWith("А")||name.toUpperCase().startsWith("СИП")){
                                answer = name + " " + name2+ "  " + String.format("%.2f", price*eaAluminiumDiscount ) + " грн/м";}
                                else {answer = name + " " + name2+ "  " + String.format("%.2f", price*eaCuprumDiscount ) + " грн/м";}

                                System.out.println(answer);
                                rezult.add(answer+ " ПРАЙС ЭНЕРГОАЛЬЯНС");
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
                            if (name.toUpperCase().contains(request.toUpperCase()) ) {
                                Double price = row.getCell(3 + 4 * j).getNumericCellValue();

                                String answer ="";
                                if(name.toUpperCase().startsWith("А")||name.toUpperCase().startsWith("СИП")){
                                    answer = name + " " + name2+ "  " + String.format("%.2f", price*eaAluminiumDiscount ) + " грн/м";}
                                else {answer = name + " " + name2+ "  " + String.format("%.2f", price*eaCuprumDiscount ) + " грн/м";}
                                System.out.println(answer);
                                rezult.add(answer + "ПРАЙС ЭНЕРГОАЛЬЯНС");
                            }
                        } catch (IllegalStateException e) {
                        }
                        catch (NullPointerException e1){
                        }
                    }
                }
            }
        }
    } //поиск в прайсе Энергоальянс


    public static void findMatchesMK(String file, String request) throws IOException{ //????? ? ???????? ?????? ??????
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row ;

        System.out.println("----------------------------------------------------------");
        System.out.println(" МАСТЕР КАБЕЛЬ");

        if(request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(1).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())) {
                        Double cuantaty = row.getCell(5).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000)+ " м ";
                        System.out.println(answer);
                        rezult.add(answer+ " МАСТЕР КАБЕЛЬ");
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

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000)+ " м ";
                        System.out.println(answer);
                        rezult.add(answer + " МАСТЕР КАБЕЛЬ");
                    }
                }
                catch (IllegalStateException e){}
                catch (NullPointerException e5){}
            }
        }
        myExcelBook.close();
    } //поиск в файле Мастер Кабель









    public static class SimpleGUI extends JFrame {
        private JButton button = new JButton("Поиск");
        public JTextField input = new JTextField("", 5);
        private JLabel label = new JLabel("Введите запрос в формате (АВВГ 3х25)");
        public JTextArea textArea = new JTextArea("");
        public JButton button2 = new JButton("Очистить");




        public SimpleGUI() {
            super("Поиск кабелей в остатках поставщиков");
            this.setBounds(500,200,300,200);
            this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
textArea.setEditable(false);
            JScrollPane jScrollPane = new JScrollPane(textArea);


            Container container = this.getContentPane();
            container.setLayout(new GridLayout(5,1,2,2));
            container.add(label,BorderLayout.NORTH);
            input.setBackground(Color.RED);
            input.setBounds(0,0,200,50);
            container.add(input);
            button.addActionListener(new ButtonEventListener());
            container.add(button);
            container.add(jScrollPane);
            button2.addActionListener(new ButtonEventListener());
            container.add(button2);



        }

        class ButtonEventListener implements ActionListener {

            public void actionPerformed(java.awt.event.ActionEvent evt) {
                if(evt.getSource()==button){
                    String str = input.getText();
                    try{
                    findMatchesDTS(dtsFileName, str);
                    findMatchesEA(eaFileName, str);
                    findMatchesInPriceEA(eaPriceFileName, str);
                    findMatchesMK(mkFileName, str);}
                    catch (IOException e){}

                    for(int i = 0; i<rezult.size();i++){
                        textArea.append(rezult.get(i)+"\n");
                    }
                }
                if(evt.getSource()==button2){
                    textArea.setText("");
                    rezult.clear();
                }
            }
        }



    }


}