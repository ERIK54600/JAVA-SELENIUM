package sample;
import com.sun.xml.internal.ws.api.pipe.FiberContextSwitchInterceptor;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
//import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.WritableSheet;

import jxl.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.File;
import java.io.IOException;
import java.util.concurrent.TimeUnit;



public class Controller {



    private String validador;

    @FXML

    private ProgressBar bar;

    @FXML
    private String file2;

    @FXML
    private String file3;

    @FXML
    private javafx.scene.control.Label label;

    @FXML
    private javafx.scene.control.Label label2;
    private String value;
    private String contraparte;
    public WebDriver driver;
    private String valueb;






    public void chooser(){
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Abrir archivo Excel");
        File file = chooser.showOpenDialog(new Stage());
        if (file != null) {
            file2=file.toString();
            label.setText(file2);
        }
    }

    public void chooserdir(){
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Seleccionar carpeta");
        File selectedDirectory =
                directoryChooser.showDialog(new Stage());
        if(selectedDirectory!=null){
            file3=selectedDirectory.toString();
            label2.setText(file3);


        }

    }




    //CLICK ON EXECUTION BUTTON
    public void click() throws  IOException,Exception {
        String check2=label2.getText();
        String check = label.getText().substring(label.getText().length() - 4);
        String valor;
        String valorb;

        valor=valida(check2,check);
        System.out.println(valor);

        if("Incorrecto".equals(value)){
            System.out.println("funciona");
            return;
        }

        System.out.println("siguio");
        File src=new File(label.getText());
        System.out.println(src);
        Workbook wb = Workbook.getWorkbook(src);
        System.out.println(src);
        Sheet sheet=wb.getSheet("Counterparties");

        WritableWorkbook wwbCopy = Workbook.createWorkbook(new File("BookTest.xls"));
        System.out.println("imprimio");
        wwbCopy.createSheet("BookTest",0);
        WritableSheet shSheet=wwbCopy.getSheet("BookTest");
        System.out.println("imprimio2");

        //Label labeler = new Label(0,0,"prueba");
        //shSheet.addCell(labeler);
        //wwbCopy.write();
        //wwbCopy.close();


        for (int i=1; i<sheet.getRows();i++){
            System.out.println(sheet.getRows());
            Cell cellcounterparty=sheet.getCell(0,i);
            String celda=cellcounterparty.getContents().toString();
            System.out.println(celda);
            //SET INTERNET EXPLORER SETTINGS
            System.setProperty("webdriver.ie.driver","C:\\Selenium\\Drivers\\IEDriver\\IEDriverServer.exe");
            DesiredCapabilities IEDesiredCapabilities = DesiredCapabilities.internetExplorer();
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.ENABLE_PERSISTENT_HOVERING, false);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.REQUIRE_WINDOW_FOCUS, false);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION, true);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.IGNORE_ZOOM_SETTING, true);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.NATIVE_EVENTS, false);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
            IEDesiredCapabilities.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, "http://www.google.com");
            IEDesiredCapabilities.internetExplorer().setCapability("ignoreProtectedModeSettings", true);
            IEDesiredCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
            IEDesiredCapabilities.setJavascriptEnabled(true);
            IEDesiredCapabilities.setCapability("ignoreZoomSetting", true);
            IEDesiredCapabilities.setCapability("enablePersistentHover", false);

            driver= new InternetExplorerDriver(IEDesiredCapabilities);
            driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
            driver.get("https://google.com.mx/");
            valorb=busquedabloomberg(celda,wwbCopy,shSheet,i);
            //wwbCopy.write();
            if ("Incorrecto".equals(valorb)){



            }





        }









    }


    public void escribelibro(){



    }

    public String leelibro(){

        return contraparte;
    }


    public String busquedabloomberg(String busquedab,WritableWorkbook wbook, WritableSheet wsheet,int k) {
        driver.findElement(By.id("lst-ib")).clear();
        driver.findElement(By.id("lst-ib")).sendKeys(busquedab + " bloomberg");
        try {
            TimeUnit.SECONDS.sleep(10);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        driver.findElement(By.name("btnK")).submit();
        java.util.List<WebElement> links=driver.findElements(By.tagName("a"));
        System.out.println(links.size());
        String data;
        WritableCell labeler;
        WritableCell labeler2;

        for (int j=1; j<links.size() ;j=j+1)
        {



            data=links.get(j).getAttribute("href");

            if(data==null){
                continue;
            }
            if(data.length()>=41) {
                data = data.substring(0, 41);


                //System.out.println(data);
                if (data.equals("https://www.bloomberg.com/research/stocks")) {
                    System.out.println(data);
                    links.get(j).click();
                    valueb="Correcto";
                    break;

                }
                else System.out.println("else");
                valueb="Incorrecto";

            }

        }

        try {
            TimeUnit.SECONDS.sleep(10);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }


        //String contenido=driver.findElement(By.id("bDesc")).getText();

        //java.util.List<WebElement> links2=driver.findElements(By.tagName("p"));
        String contenido = null;
        contenido=driver.findElement(By.xpath("//*[@id='bDesc']")).getAttribute("innerHTML");





        //System.out.println(contenido);

        if(contenido.equals("")) {

                System.out.println("vacio");
        }

        else {
            System.out.println(contenido);
            labeler = new Label(1,k,contenido);
            labeler2 = new Label(0,k,busquedab);
            try {
                wsheet.addCell(labeler);
            } catch (WriteException e) {
                e.printStackTrace();

            }

            try {
                wsheet.addCell(labeler2);
            } catch (WriteException e) {
                e.printStackTrace();
            }

            try {
                wbook.write();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                wbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }


        }

        driver.quit();
        return valueb;
    }






    public String valida(String carpeta, String archivo) {

        if(!"Seleccionar carpeta de guardado".equals(carpeta)) {

            System.out.println(archivo);
            if (archivo.equals(".xls")) {

                value="Correcto";


            }

            else if (archivo != ".xls") {
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("Error");
                alert.setHeaderText("Error en el archivo de Excel");
                alert.setContentText("El archivo de Excel no tiene el formato .xls");

                alert.showAndWait();

                value="Incorrecto";
            }



        }

        else if (carpeta.equals("Seleccionar carpeta de guardado") ) {
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Error");
            alert.setHeaderText("Error en la selección de la carpeta");
            alert.setContentText("La carpeta no se ha seleccionado correctamente");

            alert.showAndWait();

            value="Incorrecto";

        }





        return value;
    }

















    }






