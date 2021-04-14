import javafx.scene.control.TextField;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;

public class MainSceneController {

    final FileChooser fileChooser = new FileChooser();

    public TextField textPN;
    public TextField textMan;
    public TextField textAcc;
    public File BOM;
    public Text textFinish;
    public Text BOMError;
    public Text textErrorPN;
    public Text textErrorMan;
    public Text textErrorAcc;

    public void buttonClickedChoseBOM() {
        Stage stage = new Stage();
        File file = fileChooser.showOpenDialog(stage);
        if(file == null)
            return;
        BOM = new File(file.getPath());
    }

    public void buttonClickedGo() throws IOException {
        int PN;
        int Man;
        int intAcc;
        double doubleAcc;

        if (BOM == null) {
            BOMError.setText("Было бы неплохо всё таки выбрать файл...");
            return;
        }

        if (isDigit(textPN.getText())){
            PN = Integer.parseInt(textPN.getText());
        }else {
            textErrorPN.setText("Если что, сюда нужно ввести число");
            return;
        }
        if (isDigit(textMan.getText())){
            Man = Integer.parseInt(textMan.getText());
        }else {
            textErrorMan.setText("Если что, сюда нужно ввести число");
            return;
        }
        if (isDigit(textAcc.getText())){
            intAcc = Integer.parseInt(textAcc.getText());
        }else {
            textErrorAcc.setText("Если что, сюда нужно ввести число");
            return;
        }

        doubleAcc = Double.parseDouble(textAcc.getText()) / 100;


        String nameFile = BOM.getName().substring(0, BOM.getName().length() - 5);
        main.readFromExcel(BOM.getPath(), PN, Man, doubleAcc, nameFile);


        textFinish.setText("Обработанный BOM - " + nameFile + "NEW");
        BOMError.setText("");
        textErrorPN.setText("");
        textErrorMan.setText("");
        textErrorAcc.setText("");
    }

    private static boolean isDigit(String s) throws NumberFormatException {
        try {
            Integer.parseInt(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}