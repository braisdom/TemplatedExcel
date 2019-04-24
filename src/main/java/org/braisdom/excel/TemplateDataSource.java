package org.braisdom.excel;

import java.util.Set;

public interface TemplateDataSource {

    public Set<String> getDataNames();

    public Object getData(String name);

}

