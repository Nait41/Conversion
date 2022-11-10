
import data.InfoList;
import fileView.DOCOpen;
import fileView.XLSXSave;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.ResourceBundle;

public class MainController {

    public InfoList infoList;
    File notebookPath;
    DOCOpen docOpen;
    XLSXSave xlsxSave;
    File saveNotebook;
    boolean checkLoad, checkUnload, checkStart = false;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button dirUnloadButton;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Button startButton;

    @FXML
    public Button closeButton;

    public void addHinds(){

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите файл для конвертации");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите папку, в которой необходимо создать конвертированный файл");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы конвертировать файл");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть приложение");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

    }

    public void removeHinds(){
        dirLoadButton.setTooltip(null);
        dirUnloadButton.setTooltip(null);
        startButton.setTooltip(null);
        closeButton.setTooltip(null);
    }

    @FXML
    void initialize() throws IOException, InterruptedException, ClassNotFoundException {
        addHinds();
        FileInputStream loadStream = new FileInputStream(Application.rootDirPath + "\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream(Application.rootDirPath + "\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream(Application.rootDirPath + "\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);


        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus_end.setText("");
                FileChooser directoryChooser = new FileChooser();
                File file = directoryChooser.showOpenDialog(new Stage());
                notebookPath = file;
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит конвертация файла. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus_end.setText("");
                        DirectoryChooser dirChooser = new DirectoryChooser();
                        saveNotebook = dirChooser.showDialog(new Stage());
                        checkUnload = true;
                    }
                    else
                    {
                        errorMessageStr = "Происходит конвертация файла. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus_end.setText("");
                        if(checkLoad & checkUnload){
                            checkStart = true;
                            new Thread(){
                                @Override
                                public void run(){
                                    MainLoader mainLoader = null;
                                    try {
                                        mainLoader = new MainLoader(saveNotebook);
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    } catch (InvalidFormatException e) {
                                        e.printStackTrace();
                                    }
                                    if(notebookPath.getPath().contains(".docx"))
                                    {
                                        infoList = new InfoList();
                                        try {
                                            docOpen = new DOCOpen(notebookPath);
                                            docOpen.getMainInfo(infoList);
                                            docOpen.close();
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        } catch (InvalidFormatException e) {
                                            e.printStackTrace();
                                        }
                                        try {
                                            xlsxSave = new XLSXSave(Application.rootDirPath, saveNotebook, infoList);
                                            xlsxSave.setAllData();
                                            xlsxSave.saveFile();
                                            xlsxSave.close();
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        } catch (InvalidFormatException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                    loadStatus_end.setText("Файл успешно конвертирован!");
                                    checkStart = false;
                                }
                            }.start();
                        } else {
                            errorMessageStr = "Вы не указаали файл загрузки или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит конвертация файла. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
