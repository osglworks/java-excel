package org.osgl.xls;

/*-
 * #%L
 * Java Excel Reader
 * %%
 * Copyright (C) 2017 OSGL (Open Source General Library)
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

import org.osgl.util.N;
import org.osgl.util.S;

public class Address {
    private String unitNo;
    private String streetNo;
    private String street;
    private String suburb;
    private String state;
    private String postCode;

    public String getUnitNo() {
        return unitNo;
    }

    public void setUnitNo(String unitNo) {
        this.unitNo = unitNo;
    }

    public String getStreetNo() {
        return streetNo;
    }

    public void setStreetNo(String streetNo) {
        this.streetNo = streetNo;
    }

    public String getStreet() {
        return street;
    }

    public void setStreet(String street) {
        this.street = street;
    }

    public String getSuburb() {
        return suburb;
    }

    public void setSuburb(String suburb) {
        this.suburb = suburb;
    }

    public String getState() {
        return state;
    }

    public void setState(String state) {
        this.state = state;
    }

    public String getPostCode() {
        return postCode;
    }

    public void setPostCode(String postCode) {
        this.postCode = postCode;
    }

    public static Address random() {
        Address address = new Address();
        address.setPostCode(S.string(N.randInt(8999) + 1000));
        address.setState(S.random(3) + 1);
        address.setStreet(S.random(N.randInt(20) + 10) );
        address.setSuburb(S.random(N.randInt(8) + 5));
        address.setStreetNo(S.string(N.randInt(100)));
        address.setUnitNo(S.string(N.randInt(10)));
        return address;
    }
}
