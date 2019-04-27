/**
 * Copyright (C) 2019-2025 Braisdom Wang (www.joowing.com)
 * wangyonghe@msn.com
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.braisdom.excel;

import java.util.Set;

/**
 * Defining a context variables for the template, it will be used for
 * <a href="https://github.com/thymeleaf/thymeleaf">thymeleaf</a>.
 *
 * @author braisdom
 * @since 1.0.0
 */
public interface TemplateDataSource {

    /**
     * Returns a list with all the names of variables contained at this context.
     *
     * @return the variable names.
     */
    public Set<String> getDataNames();

    /**
     * Retrieve a specific variable, by name.
     * @param name the name of the variable to be retrieved.
     * @return the variable's value.
     */
    public Object getData(String name);

}

