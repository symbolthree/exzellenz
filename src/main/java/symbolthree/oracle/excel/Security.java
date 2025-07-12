/******************************************************************************
 *
 * ≡ EXZELLENZ ≡
 * Copyright (C) 2009-2022 Christopher Ho 
 * All Rights Reserved, http://www.symbolthree.com
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 *
 * E-mail: christopher.ho@symbolthree.com
 *
******************************************************************************/

package symbolthree.oracle.excel;

import java.lang.reflect.Method;

public class Security implements Constants {
    private static Security security      = null;
    private String          decryptMethod = "";
    private String          encryptMethod = "";
    private String          securityClass = "";
    
    protected Security() {
		Class[] params = new Class[3];

        params[0] = String.class;
        params[1] = Integer.TYPE;
        params[2] = String.class;

        try {
            Class<?> secClass = Class.forName("oracle.apps.fnd.security.AolSecurity");

            secClass.getDeclaredMethod("encrypt", params);
            securityClass = "oracle.apps.fnd.security.AolSecurity";
            encryptMethod = "encrypt";
            decryptMethod = "decrypt";
        } catch (Exception e) {
        }

        if (securityClass.equals("")) {
            try {
                Class<?> secClass = Class.forName("oracle.apps.fnd.security.Security");

                secClass.getDeclaredMethod("encrypt", params);
                securityClass = "oracle.apps.fnd.security.Security";
                encryptMethod = "encrypt";
                decryptMethod = "decrypt";
            } catch (Exception e) {
            }
        }

        if (securityClass.equals("")) {
            try {
                Class<?> secClass = Class.forName("oracle.apps.fnd.security.Security");

                secClass.getDeclaredMethod("f", params);
                securityClass = "oracle.apps.fnd.security.Security";
                encryptMethod = "f";
                decryptMethod = "k";
            } catch (Exception e) {
            	EXZHelper.log(LOG_DEBUG, "oracle.apps.fnd.security.Security (fndext) not found or incorrect method");
            }
        }
    }

    public static Security getInstance() {
        if (security == null) {
            security = new Security();
        }

        return security;
    }

    
    protected boolean isEncrypted(String str) {
    	if (str.startsWith("ZG") && str.length()==ENCRYPTED_PASSWORD_LENGTH) {
    		return true;
    	} else {
    		return false;
    	}
    }
    
    protected String encryptPwd(String str) {
    	if (isEncrypted(str)) return str;

        String pwd = str;

        try {
            Class<?>   secClass = Class.forName(securityClass);
            Object  secObj   = secClass.getDeclaredConstructor().newInstance();
            Class[] params   = new Class[3];

            params[0] = String.class;
            params[1] = Integer.TYPE;
            params[2] = String.class;

            Object[] objs   = { PASSWORD_SEED, Integer.valueOf(ENCRYPTED_PASSWORD_LENGTH), str };
            Method   method = secClass.getDeclaredMethod(encryptMethod, params);

            pwd = (String) method.invoke(secObj, objs);
        } catch (Exception e) {
        	EXZHelper.log(LOG_ERROR, e.getMessage());
        }

        return pwd;
    }

    protected String decryptPwd(String hash) {
    	if (! isEncrypted(hash)) return hash;
    	
        String pwd = hash;
        
        try {
            Class<?>   secClass = Class.forName(securityClass);
            Object  secObj   = secClass.getDeclaredConstructor().newInstance();
            Class[] params   = new Class[2];

            params[0] = String.class;
            params[1] = String.class;

            Object[] objs   = { PASSWORD_SEED, hash };
            Method   method = secClass.getDeclaredMethod(decryptMethod, params);

            pwd = (String) method.invoke(secObj, objs);
        } catch (Exception e) {
        	EXZHelper.log(LOG_ERROR, e.getMessage());
        }

        return pwd;
    }
}

