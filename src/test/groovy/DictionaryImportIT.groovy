import io.dbmaster.testng.BaseToolTestNGCase;
import io.dbmaster.testng.OverridePropertyNames;
import static org.testng.Assert.assertTrue;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test

@OverridePropertyNames(project="project.dictionary")
public class DictionaryImportIT extends BaseToolTestNGCase {

    @Test
    @Parameters(["dictionary-import.p_excel_file","dictionary-import.p_model",
                 "dictionary-import.p_object_tab", "dictionary-import.p_column_tab",
                 "dictionary-import.p_field_mapping"])
    public void importDictionary(String p_excel_file, String p_model, 
        String p_object_tab,String p_column_tab, String p_field_mapping) {
        def parameters = [ "p_excel_file"  :  p_excel_file,
                           "p_model"       :  p_model,
                           "p_object_tab"  :  p_object_tab,
                           "p_column_tab"  :  p_column_tab,
                           "p_mode"        :  "Sync",
                           "p_action"      :  "Preview",
                           "p_field_mapping" : p_field_mapping]
        tools.toolExecutor("dictionary-import", parameters).execute()
    }
}
