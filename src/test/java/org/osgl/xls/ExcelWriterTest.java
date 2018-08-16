package org.osgl.xls;

/*-
 * #%L
 * Java Excel Tool
 * %%
 * Copyright (C) 2017 - 2018 OSGL (Open Source General Library)
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

import org.junit.Before;
import org.junit.Test;
import org.osgl.util.S;
import osgl.ut.TestBase;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriterTest extends TestBase {

    private File targetFile;

    private List<Student> students;

    @Before
    public void prepare() throws IOException {
        targetFile = File.createTempFile("student", ".xlsx");
        students = new ArrayList<>();
        for (int i = 0; i < 10; ++i) {
            students.add(Student.random(i));
        }
    }

    @Test
    public void test() {
        ExcelWriter writer = ExcelWriter.builder()
                .dateFormat("dd/MMM/yyyy")
                .headerTransformer(S.F.dropHeadIfStartsWith("address."))
                .filter("-address,+address.postCode")
                .build();
        writer.write(students, targetFile);
    }

}
