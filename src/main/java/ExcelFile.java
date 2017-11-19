import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public interface ExcelFile {
    public ArrayList<String> findMatches(String request) throws IOException;
}
