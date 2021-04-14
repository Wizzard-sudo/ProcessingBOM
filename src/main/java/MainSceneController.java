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

    public void buttonClickedChoseBOM() {
        Stage stage = new Stage();
        File file = fileChooser.showOpenDialog(stage);
        if(file == null)
            return;
        BOM = new File(file.getPath());
    }

    public void buttonClickedGo() throws IOException {

        if (BOM == null) {
            BOMError.setText("Было бы неплохо всё таки выбрать файл...");
            return;
        }


        String nameFile = BOM.getName().substring(0, BOM.getName().length() - 5);
        main.readFromExcel(BOM.getPath(),
                Integer.parseInt(textPN.getText()),
                Integer.parseInt(textMan.getText()),
                Double.parseDouble(textAcc.getText().replaceAll(",", ".")),
                nameFile);


        textFinish.setText("Обработанный BOM - " + nameFile + "NEW");
    }
}
