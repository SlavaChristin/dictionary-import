import java.io.InputStream;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;




import org.apache.commons.io.IOUtils
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.usermodel.DateUtil
import org.slf4j.Logger;




import com.branegy.dbmaster.database.api.ModelService
import com.branegy.dbmaster.model.Model;
import com.branegy.dbmaster.model.ModelObject;
import com.branegy.dbmaster.util.NameMap;

import javax.persistence.EntityManager;

import com.branegy.util.InjectorUtil;



import com.branegy.dbmaster.model.Column;
import com.branegy.dbmaster.model.Parameter;
import com.branegy.dbmaster.model.Table;
import com.branegy.dbmaster.model.View;;
import com.branegy.dbmaster.model.Procedure;
import com.branegy.dbmaster.model.Function;

import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;

import com.branegy.dbmaster.custom.CustomFieldConfig;
import com.branegy.dbmaster.custom.field.server.api.ICustomFieldService;
import com.branegy.util.InjectorUtil;


def findColumn(List<Column> list,String name){
    if (list != null){
        for (Column c:list){
            if (c.getName().equals(name)){
                return c;
            }
        } 
    }
    return null;
}


def sync(Sheet objectTab, Sheet columnTab, Model source, Model target, Map<String,Set<String>> customFields){
    if (objectTab == null){
        logger.error("Object tab is required");
        return;
    }
    if (columnTab == null){
        logger.error("Column tab is required");
        return;
    }
    
    ExcelReaderIterator objectIt = new ExcelReaderIterator(objectTab, p_field_mapping,
        ["Type":true, "Name":true],
        [:],
        ["Table","View","Procedure","Function"] as Set,
        logger
    );
    ExcelReaderIterator columnIt = new ExcelReaderIterator(columnTab, p_field_mapping,
        ["Object":true, "Name":true, "Type":true, "Nullable": false,"Size":false,"Precision":false,"Scale":false],
        ["Default value":false, "Definition": false], // param type
        ["Column","Parameter"] as Set,
        logger
    );

    customFields.put("Table", objectIt.getCustomNames("Table"));
    customFields.put("View", objectIt.getCustomNames("View"));
    customFields.put("Procedure", objectIt.getCustomNames("Procedure"));
    customFields.put("Function", objectIt.getCustomNames("Function"));
    
    customFields.put("Column", columnIt.getCustomNames("Column"));
    customFields.put("Parameter", columnIt.getCustomNames("Parameter"));

    NameMap<ModelObject> existsObjects = new NameMap<ModelObject>();
    NameMap<List<Column>> existsObjectColumns = new NameMap<List<Column>>()
    target.getTables().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
        it.setColumns(new ArrayList<Column>());
    }
    target.setTables(new ArrayList<Table>());
    target.getViews().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
        it.setColumns(new ArrayList<Column>());
    }
    target.setViews(new ArrayList<View>());
    target.getFunctions().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
        it.setParameters(new ArrayList<Parameter>());
    }
    target.setFunctions(new ArrayList<Function>());
    target.getProcedures().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
        it.setParameters(new ArrayList<Parameter>());
    }
    target.setProcedures(new ArrayList<Procedure>());

    NameMap<ModelObject> objects = new NameMap<ModelObject>();
    while (objectIt.nextRow()){
        try{
            String type = StringUtils.capitalize(objectIt.getColumnString(0).toLowerCase());
            String name = objectIt.getColumnString(1);
            
            ModelObject object = existsObjects.get(name);
            if (object == null
                 || !org.hibernate.Hibernate.getClass(object).getClass().getSimpleName().equals(type)){
                switch(type){
                    case "Table":
                        object = new Table();
                        object.setColumns(new ArrayList<Column>());
                        break;
                    case "View":
                        object = new View();
                        object.setColumns(new ArrayList<Column>());
                        break;
                    case "Procedure":
                        object = new Procedure();
                        object.setParameters(new ArrayList<Parameter>());
                        break;
                    case "Function":
                        object = new Function();
                        object.setParameters(new ArrayList<Parameter>());
                        break;
                    default:
                        logger.warn("Unknown object type {}, should be {}", type, "[Table,View,Procedure,Function]");
                        continue;
                }
                object.setName(name);
            }

            object.getCustomMap().putAll(objectIt.getCustomValues(type));
            
            if (objects.put(name, object) != null){
                logger.error("Duplicate object {}", name);
            }   
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading objects: {}",e.getMessage());
            continue;
        }
    }
    
    while (columnIt.nextRow()){
        try{
            String parent = columnIt.getColumnString(0);
            String name =   columnIt.getColumnString(1);
            String type =   columnIt.getColumnString(2);
            boolean nullable = Boolean.TRUE.equals(columnIt.getColumnBoolean(3));
            Integer size =  columnIt.getColumnInteger(4);
            Integer precision = columnIt.getColumnInteger(5);
            Integer scale =  columnIt.getColumnInteger(6);
            
            if (!objects.containsKey(parent)){
                logger.error("Parent object {} is not found at row {}" , parent,columnIt.getRowNumber());
                continue;
            }
            
            ModelObject object = objects.get(parent);
            Column col;
            if ((object instanceof View) || (object instanceof Table)){
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Column();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Column"));
                object.addColumn(col);
            } else {
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Parameter();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Parameter"));
                object.addParameter(col);
            }
            
            col.setType(type);
            col.setNullable(nullable);
            col.setSize(size);
            col.setPrecesion(precision);
            col.setScale(scale);
            
            if (columnIt.hasColumn(7)){
                col.setDefaultValue(columnIt.getColumnString(7));
            }
            if (columnIt.hasColumn(8)){
                col.setExtraDefinition(columnIt.getColumnString(8));
            }
           
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading columns: {}",e.getMessage());
            continue;
        }
    }
   
    objects.values().each { 
        if (it instanceof Table){
            target.addTable(it);
        } else if (it instanceof View){
            target.addView(it);
        } else if (it instanceof Procedure){
            target.addProcedure(it);
        } else {
            target.addFunction(it);
        }
    }
}

def imp(Sheet objectTab, Sheet columnTab, Model source, Model target, Map<String,Set<String>> customFields){
    ExcelReaderIterator objectIt;
    ExcelReaderIterator columnIt;
    if (objectTab == null){
        objectIt = new ExcelReaderIterator();
    } else {
        objectIt = new ExcelReaderIterator(objectTab, p_field_mapping,
            ["Type":true, "Name":true],
            [:],
            ["Table","View","Procedure","Function"] as Set,
            logger
        );
        customFields.put("Table", objectIt.getCustomNames("Table"));
        customFields.put("View", objectIt.getCustomNames("View"));
        customFields.put("Procedure", objectIt.getCustomNames("Procedure"));
        customFields.put("Function", objectIt.getCustomNames("Function"));
    }
    if (columnTab == null){
        objectIt = new ExcelReaderIterator();
    } else {
        columnIt = new ExcelReaderIterator(columnTab, p_field_mapping,
            ["Object":true, "Name":true, "Type":true, "Nullable": false,"Size":false,"Precision":false,"Scale":false],
            ["Default value":false, "Definition": false], // param type
            ["Column","Parameter"] as Set,
            logger
        );
        customFields.put("Column", columnIt.getCustomNames("Column"));
        customFields.put("Parameter", columnIt.getCustomNames("Parameter"));
    }
    
    NameMap<ModelObject> existsObjects = new NameMap<ModelObject>();
    NameMap<List<Column>> existsObjectColumns = new NameMap<List<Column>>()
    target.getTables().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
        it.setColumns(new ArrayList<Column>());
    }
    target.setTables(new ArrayList<Table>());
    target.getViews().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
        it.setColumns(new ArrayList<Column>());
    }
    target.setViews(new ArrayList<View>());
    target.getFunctions().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
        it.setParameters(new ArrayList<Parameter>());
    }
    target.setFunctions(new ArrayList<Function>());
    target.getProcedures().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
        it.setParameters(new ArrayList<Parameter>());
    }
    target.setProcedures(new ArrayList<Procedure>());

    NameMap<ModelObject> objects = new NameMap<ModelObject>();
    while (objectIt.nextRow()){
        try{
            String type = StringUtils.capitalize(objectIt.getColumnString(0).toLowerCase());
            String name = objectIt.getColumnString(1);
            
            ModelObject object = existsObjects.get(name);
            if (object == null
                 || !org.hibernate.Hibernate.getClass(object).getClass().getSimpleName().equals(type)){
                switch(type){
                    case "Table":
                        object = new Table();
                        object.setColumns(new ArrayList<Column>());
                        break;
                    case "View":
                        object = new View();
                        object.setColumns(new ArrayList<Column>());
                        break;
                    case "Procedure":
                        object = new Procedure();
                        object.setParameters(new ArrayList<Parameter>());
                        break;
                    case "Function":
                        object = new Function();
                        object.setParameters(new ArrayList<Parameter>());
                        break;
                    default:
                        logger.warn("Unknown object type {}, should be {}", type, "[Table,View,Procedure,Function]");
                        continue;
                }
                object.setName(name);
            }

            object.getCustomMap().putAll(objectIt.getCustomValues(type));
            
            if (objects.put(name, object) != null){
                logger.error("Duplicate object {}", name);
            }
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading objects: {}",e.getMessage());
            continue;
        }
    }
    
    while (columnIt.nextRow()){
        try{
            String parent = columnIt.getColumnString(0);
            String name =   columnIt.getColumnString(1);
            String type =   columnIt.getColumnString(2);
            boolean nullable = Boolean.TRUE.equals(columnIt.getColumnBoolean(3));
            Integer size =  columnIt.getColumnInteger(4);
            Integer precision = columnIt.getColumnInteger(5);
            Integer scale =  columnIt.getColumnInteger(6);
            
            if (!objects.containsKey(parent)){
                logger.error("Parent object {} is not found at row {}" , parent,columnIt.getRowNumber());
                continue;
            }
            
            ModelObject object = objects.get(parent);
            Column col;
            if ((object instanceof View) || (object instanceof Table)){
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Column();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Column"));
                object.addColumn(col);
            } else {
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Parameter();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Parameter"));
                object.addParameter(col);
            }
            
            col.setType(type);
            col.setNullable(nullable);
            col.setSize(size);
            col.setPrecesion(precision);
            col.setScale(scale);
            
            if (columnIt.hasColumn(7)){
                col.setDefaultValue(columnIt.getColumnString(7));
            }
            if (columnIt.hasColumn(8)){
                col.setExtraDefinition(columnIt.getColumnString(8));
            }
           
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading columns: {}",e.getMessage());
            continue;
        }
    }
   
    objects.values().each {
        if (it instanceof Table){
            target.addTable(it);
        } else if (it instanceof View){
            target.addView(it);
        } else if (it instanceof Procedure){
            target.addProcedure(it);
        } else {
            target.addFunction(it);
        }
    }
    existsObjects.values().each { object ->
        if (!objects.containsKey(object.getName())){
            if ((object instanceof View) || (object instanceof Table)){
                object.setColumns(existsObjectColumns.get(object.getName()));
            } else {
                object.setParameters(existsObjectColumns.get(object.getName()));
            }
        }
    }
}

def met(Sheet objectTab, Sheet columnTab, Model source, Model target, Map<String,Set<String>> customFields){
    ExcelReaderIterator objectIt;
    ExcelReaderIterator columnIt;
    if (objectTab == null){
        objectIt = new ExcelReaderIterator();
    } else {
        objectIt = new ExcelReaderIterator(objectTab, p_field_mapping,
            ["Type":true, "Name":true],
            [:],
            ["Table","View","Procedure","Function"] as Set,
            logger
        );
    
        customFields.put("Table", objectIt.getCustomNames("Table"));
        customFields.put("View", objectIt.getCustomNames("View"));
        customFields.put("Procedure", objectIt.getCustomNames("Procedure"));
        customFields.put("Function", objectIt.getCustomNames("Function"));
    }
    if (columnTab == null){
        objectIt = new ExcelReaderIterator();
    } else {
        columnIt = new ExcelReaderIterator(columnTab, p_field_mapping,
            ["Object":true, "Name":true, "Type":false, "Nullable": false,"Size":false,"Precision":false,"Scale":false],
            ["Default value":false, "Definition": false], // param type
            ["Column","Parameter"] as Set,
            logger
        );

        customFields.put("Column", columnIt.getCustomNames("Column"));
        customFields.put("Parameter", columnIt.getCustomNames("Parameter"));
    }
    
    NameMap<ModelObject> existsObjects = new NameMap<ModelObject>();
    NameMap<List<Column>> existsObjectColumns = new NameMap<List<Column>>()
    target.getTables().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
    }
    target.getViews().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getColumns())
    }
    target.getFunctions().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
    }
    target.getProcedures().each {
        existsObjects.put(it.getName(), it)
        existsObjectColumns.put(it.getName(), it.getParameters())
    }

    while (objectIt.nextRow()){
        try{
            String type = StringUtils.capitalize(objectIt.getColumnString(0).toLowerCase());
            String name = objectIt.getColumnString(1);
            ModelObject object = existsObjects.get(name);
            if (object == null){
                continue;
            }
            object.getCustomMap().putAll(objectIt.getCustomValues(type));
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading objects: {}",e.getMessage());
            continue;
        }
    }
    
    while (columnIt.nextRow()){
        try{
            String parent = columnIt.getColumnString(0);
            String name =   columnIt.getColumnString(1);
            
            if (!existsObjects.containsKey(parent)){
                logger.error("Parent object {} is not found at row {}" , parent,columnIt.getRowNumber());
                continue;
            }
            
            ModelObject object = existsObjects.get(parent);
            Column col;
            if ((object instanceof View) || (object instanceof Table)){
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Column();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Column"));
            } else {
                col = findColumn(existsObjectColumns.get(parent), name);
                if (col == null){
                    col = new Parameter();
                    col.setName(name);
                }
                col.getCustomMap().putAll(columnIt.getCustomValues("Parameter"));
            }
        } catch (IllegalArgumentException e){
            logger.warn("Error while reading columns: {}",e.getMessage());
            continue;
        }
    }
}

InputStream fis = null
try {
    boolean error = false;
    fis = p_excel_file.getInputStream()
    Workbook wb = WorkbookFactory.create(fis)
    
    Sheet objectTab = p_object_tab!=null?wb.getSheet(p_object_tab):null;
    Sheet columnTab = p_column_tab!=null?wb.getSheet(p_column_tab):null;
    if (p_object_tab!=null && objectTab == null){
        logger.error("Sheet {} is not found", p_object_tab);
        error = true;
    }
    if (p_column_tab!=null && columnTab == null){
        logger.error("Sheet {} is not found", p_column_tab);
        error = true;
    }

    ModelService service = dbm.getService(ModelService.class);
    
    Model target = service.findModelByName(p_model.split("\\.",2)[0],p_model.split("\\.",2)[1], Model.FETCH_TREE);
    InjectorUtil.getInstance(EntityManager.class).detach(target);
    Model source = service.findModelById(target.getId(), Model.FETCH_TREE);
    
    Map<String,Set<String>> customFields = new HashMap<String,Set<String>>();
    
    if (!error){
        switch (p_mode){
            case "Sync":
                sync(objectTab, columnTab, source, target, customFields);
                break;
            case "Import":
                imp(objectTab, columnTab, source, target, customFields);
                break;
            case "Import metadata only":
                met(objectTab, columnTab, source, target, customFields);
                break;
        }
        def syncSession = service.compareObjects(source,target, Collections.singletonMap("customFieldMap", customFields));
        com.branegy.dbmaster.sync.api.SyncService s = InjectorUtil.getInstance(com.branegy.dbmaster.sync.api.SyncService.class);
        println s.generateSyncSessionPreviewHtml("/preview-model-generator.groovy",syncSession, false);
        if ("Import".equals(p_action)){
            syncSession.applyChanges();
        }
    }
} catch (Exception e) {
    dbm.setRollbackOnly();
    throw e;
} finally {
    IOUtils.closeQuietly(fis);
}











public class ExcelReaderIterator{
    private final Iterator<Row> it;
    private final int[] required;
    private final int[] optional;
    private final boolean[] requiredValue;
    private final Map<String,Map<Integer,String>> typeColumns;
    private final Map<String,Integer> idIndex;
    
    private Row current;
    
    
    /**
     * create empty ExcelReaderIterator
     */
    public ExcelReaderIterator(){
        it = Collections.emptyList().iterator();
        required = new int[0];
        optional = new int[0];
        requiredValue = new boolean[0];
        typeColumns = Collections.emptyMap();
        idIndex = Collections.emptyMap();
    }
    
    /**
     * mappingText:
     * key=value                  \r\n | \n
     * key=                       skip field
     *
     * requiredColumns            name + required
     * optionalColumns            name + required
     * customObjectTypes          type1,...,typeN
     */
    public ExcelReaderIterator(Sheet sheet, String mappingText,
            Map<String, Boolean> requiredColumns, // name + required
            Map<String, Boolean> optionalColumns, // name + required
            Set<String> customObjectTypes,        // type1,...,typeN
            Logger logger){
        Map<String,String> mapping = new HashMap<String,String>();
        if (mappingText!=null && !mappingText.isEmpty()){
            for (String row:mappingText.split("(\r\n|\n)")){
                if (!row.contains('=') || row.trim().isEmpty()){
                    continue;
                }
                String[] kv = row.split("=",2);
                mapping.put(kv[0], kv[1].isEmpty()?null:kv[1]);
            }
        }
        it = sheet.rowIterator();
        Row header = it.next();
        
        String sheetName = sheet.getSheetName();
        idIndex = new LinkedHashMap<String, Integer>();
        for (Cell cell:header){
            if (cell.getCellType() == Cell.CELL_TYPE_STRING){
                String name = cell.getStringCellValue();
                if (name.isEmpty()){
                    continue;
                }
                if (mapping.containsKey(sheetName+"."+name)){
                    name = mapping.get(sheetName+"."+name);
                }
                
                logger.debug("Sheet {}, mapping column {} &rarr; {}",sheetName, cell.getStringCellValue(),
                    name == null? "&lt;skipping&gt;" : name);
                
                if (name!=null && idIndex.put(name, cell.getColumnIndex()) != null){
                    throw new IllegalArgumentException("Column "+name+" already exist for sheet "+sheetName);
                }
            }
        }
        if (!idIndex.keySet().containsAll(requiredColumns.keySet())){
            requiredColumns.keySet().removeAll(idIndex.keySet());
            throw new IllegalArgumentException("Required columns "+requiredColumns.keySet()
                +" is not found in "+sheetName);
        }
        requiredValue = new boolean[requiredColumns.size()+optionalColumns.size()];
        
        int i;
        required = new int[requiredColumns.size()];
        i = 0;
        for (Entry<String,Boolean> e:requiredColumns.entrySet()){
            required[i] = idIndex.remove(e.getKey());
            requiredValue[i] = e.getValue();
            i++;
        }
        optional = new int[optionalColumns.size()];
        i = 0;
        for (Entry<String,Boolean> e:optionalColumns.entrySet()){
            if (idIndex.containsKey(e.getKey())){
                optional[i] = idIndex.remove(e.getKey());
            } else {
                optional[i] = -1;
            }
            requiredValue[required.length+i] = e.getValue();
            i++;
        }
        
        // validate custom fields
        if (!customObjectTypes.isEmpty()){
            if (customObjectTypes.size() == 1){
                typeColumns = Collections.<String,Map<Integer, String>>singletonMap(customObjectTypes.iterator().next(),
                        new HashMap<Integer, String>());
            } else {
                typeColumns = new HashMap<String, Map<Integer,String>>();
                for (String type:customObjectTypes){
                    typeColumns.put(type, new HashMap<Integer, String>());
                }
            }
            ICustomFieldService service = InjectorUtil.getInstance(ICustomFieldService.class);
            List<CustomFieldConfig> configList = service.getProjectCustomConfigList();
            // check type
            for (CustomFieldConfig c:configList){
                if (customObjectTypes.contains(c.getClazz()) && idIndex.containsKey(c.getName())){
                    typeColumns.get(c.getClazz()).put(idIndex.get(c.getName()), c.getName());
                }
            }
        } else {
            typeColumns = Collections.emptyMap();
        }
    }
    
    public boolean nextRow(){
        if (it.hasNext()){
            current = it.next();
            return true;
        } else {
            return false;
        }
    }
    
    public Object getColumnValue(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readRaw(required[index],requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readRaw(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public String getColumnString(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readString(required[index], requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readString(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public Double getColumnDouble(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readDouble(required[index], requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readDouble(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public Long getColumnLong(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readLong(required[index], requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readLong(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public Integer getColumnInteger(int index){
        Long result = getColumnLong(index);
        return result!=null ? Integer.valueOf(result.intValue()):null;
    }
    
    public Date getColumnDate(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readDate(required[index], requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readDate(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public Boolean getColumnBoolean(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return readBoolean(required[index], requiredValue[index]);
        }
        index -= required.length;
        if (index < optional.length){
            if (optional[index] == -1){
                throw new IllegalArgumentException("Optional column "+index+" is not present");
            }
            return readBoolean(optional[index], requiredValue[required.length+index]);
        }
        throw new IndexOutOfBoundsException();
    }
    
    public boolean hasColumn(int index){
        if (index < 0){
            throw new IndexOutOfBoundsException();
        }
        if (index < required.length){
            return true;
        }
        index -= required.length;
        if (index < optional.length){
            return optional[index] != -1;
        }
        return false;
    }
    
    public int getRowNumber(){
        return current == null? 1: current.getRowNum()+1;
    }
    
    public Set<String> getCustomNames(String clazz){
        Map<Integer, String> map = typeColumns.get(clazz);
        if (map == null){
            throw new IllegalArgumentException("Unknown type "+clazz+", known "+typeColumns.keySet());
        }
        if (map == null){
            throw new IllegalArgumentException("Unknown type "+clazz+", known "+typeColumns.keySet());
        }
        return new LinkedHashSet<String>(map.values());
    }
    
    public Map<String,Object> getCustomValues(String clazz){
        Map<Integer, String> map = typeColumns.get(clazz);
        if (map == null){
            throw new IllegalArgumentException("Unknown type "+clazz+", known "+typeColumns.keySet());
        }
        Map<String,Object> result = new HashMap<String, Object>(map.size());
        for (Entry<Integer,String> e:map.entrySet()){
            result.put(e.getValue(), readRaw(e.getKey(),false));
        }
        return result;
    }
    
    public Set<String> getColumnNames(){
        return new LinkedHashSet<String>();
    }
    
    public Map<String,Object> getColumnValues(){
        Map<String,Object> result = new HashMap<String, Object>(idIndex.size());
        for (Entry<String,Integer> e:idIndex.entrySet()){
            result.put(e.getKey(), readRaw(e.getValue(), false));
        }
        return result;
    }
    
    private <T> T checkRequired(T result, int index, boolean required){
        if (required && (result == null || "".equals(result))){
            throw new IllegalArgumentException("Value is required for "+getRowNumber()+":"+(index+1));
        }
        return result;
    }

    private Object readRaw(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                throw new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    
    
    private String readString(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue()?"Yes":"No";
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue()+"";
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                throw new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    private Boolean readBoolean(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue().toLowerCase().matches("(true|yes)");
            default:
                throw new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    private Date readDate(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getDateCellValue();
            default:
                throw new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    private Double readDouble(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return Double.parseDouble(cell.getStringCellValue());
            default:
                tthrow new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    private Long readLong(int index){
        Cell cell = current.getCell(index);
        if (cell == null){
            return null;
        } else {
            switch (cell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_NUMERIC:
                return (long)cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return Long.parseLong(cell.getStringCellValue());
            default:
                throw new IllegalArgumentException("Unsupported cell type "+getRowNumber()+":"+(index+1));
            }
        }
    }
    
    private Object readRaw(int index, boolean required){
        return checkRequired(readRaw(index), index, required);
    }
    
    private String readString(int index, boolean required){
        return checkRequired(readString(index), index, required);
    }
    
    private Boolean readBoolean(int index, boolean required){
        return checkRequired(readBoolean(index), index, required);
    }
    
    private Double readDouble(int index, boolean required){
        return checkRequired(readDouble(index), index, required);
    }
    
    private Long readLong(int index, boolean required){
        return checkRequired(readLong(index), index, required);
    }
    
    private Date readDate(int index, boolean required){
        return checkRequired(readDate(index), index, required);
    }

    
}

