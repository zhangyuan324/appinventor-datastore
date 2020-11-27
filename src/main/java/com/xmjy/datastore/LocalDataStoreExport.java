package com.xmjy.datastore;

import com.google.appengine.api.datastore.*;
import com.xmjy.constant.CommonConstant;
import org.apache.commons.lang3.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.*;

public class LocalDataStoreExport {

    private final DatastoreService ds;

    public  LocalDataStoreExport() {
        ds = DatastoreServiceFactory.getDatastoreService();
    }

    /**
     * 按索引导出数据到csv
     * @param kind
     * @param resp
     * @throws IOException
     */
    public void exportDataToCsv(String kind, HttpServletResponse resp) throws IOException {
        List<String> kindList = getKindsList();
        if (StringUtils.isNotBlank(kind)) {
            if (!kindList.contains(kind)) {
                resp.getWriter().write("kind doesn't exist.");
                return;
            }
            kindList = Arrays.asList(kind);
        }
        StringBuffer data = new StringBuffer();
        for (String currentkind : kindList) {
            data.append(getKindData(currentkind));
            data.append("\n");
        }
        resp.setCharacterEncoding("utf-8");
        resp.setHeader("content-disposition",
                "attachment;filename=" + URLEncoder.encode(kind, "UTF-8") + ".csv");
        OutputStream outputStream = resp.getOutputStream();
        outputStream.write(data.toString().getBytes());
        outputStream.close();
    }

    /**
     * 查询实体数据
     * @param kind
     * @return
     */
    private String getKindData(String kind) {
        HashSet<String> propertyNames = new HashSet<>();
        ArrayList<Entity> entities = new ArrayList<>();

        FetchOptions fetchOptions = FetchOptions.Builder.withPrefetchSize(CommonConstant.PrefetchSize).chunkSize(CommonConstant.ChunkSize);
        Query query = new Query(kind);

        for (Entity entity : ds.prepare(query).asIterable(fetchOptions)) {
            entities.add(entity);
            propertyNames.addAll(entity.getProperties().keySet());
        }
        StringBuffer data = new StringBuffer();

        setHeaderRow(data, propertyNames);

        for (Entity entity : entities) {
            setEntity(data, entity, propertyNames);
        }
        return data.toString();
    }

    private void setHeaderRow(StringBuffer data, HashSet<String> propertyNames) {
        data.append(Entity.KEY_RESERVED_PROPERTY);
        Iterator<String> iterator = propertyNames.iterator();
        if (iterator.hasNext()) {
            data.append(CommonConstant.FieldSeparator);
        }
        while (iterator.hasNext()) {
            data.append(iterator.next());
            if (iterator.hasNext())
                data.append(CommonConstant.FieldSeparator);
        }
        data.append("\n");
    }

    private void setEntity(StringBuffer data, Entity entity, HashSet<String> propertyNames) {
        data.append(KeyFactory.keyToString(entity.getKey()));
        Iterator<String> iterator = propertyNames.iterator();
        if (iterator.hasNext()) {
            data.append(CommonConstant.FieldSeparator);
            while (iterator.hasNext()) {
                Object value = entity.getProperty(iterator.next());
                data.append(convertObjectToStr(value));
                if (iterator.hasNext())
                    data.append(CommonConstant.FieldSeparator);
            }
        }
        data.append("\n");
    }

    private String convertObjectToStr(Object value) {
        if( value == null )
            return null;

        else if( value instanceof Key )
            return KeyFactory.keyToString((Key)value);

        else if( value instanceof Date )
            return Long.toString(((Date)value).getTime());

        else if( value instanceof Enum )
            return ((Enum<?>)value).name();

        else if( value instanceof Text )
            return ((Text)value).getValue();

        else
            return value.toString();
    }

    public List<String> getKindsList() {
        ArrayList<String> kindsList = new ArrayList<>();
        Query query = new Query(Entities.KIND_METADATA_KIND);
        FetchOptions  fetchOptions = FetchOptions.Builder.withPrefetchSize(CommonConstant.PrefetchSize);
        for (Entity entity : ds.prepare(query).asIterable(fetchOptions)) {
            kindsList.add(entity.getKey().getName());
        }
        return kindsList;
    }

}
