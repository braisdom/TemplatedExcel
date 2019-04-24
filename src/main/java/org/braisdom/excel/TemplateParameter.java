package org.braisdom.excel;

import java.util.Set;

public interface TemplateParameter {

    public String getTemplateContent();

    public DataQueryParameter getDataQueryParameter(String name);

    public Set<String> getDataSourceNames();
}
