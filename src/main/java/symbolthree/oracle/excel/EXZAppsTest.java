package symbolthree.oracle.excel;

/*
import org.apache.commons.httpclient.Cookie;
import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.NameValuePair;
import org.apache.commons.httpclient.methods.PostMethod;
*/
public class EXZAppsTest {
    String instanceName = "symplik";
    String loginURL     = "http://vm01FU5010.symplik.com:8000/OA_HTML/fndvald.jsp";

    public static void main(String[] args) {
        EXZAppsTest test = new EXZAppsTest();

        // test.run();
    }

/*
    private void run() {
        try {
            HttpClient      client = new HttpClient();
            PostMethod      post   = new PostMethod(loginURL);
            NameValuePair[] data   = { new NameValuePair("username", "strengthen"),
                                       new NameValuePair("password", "strengthen") };

            post.setRequestBody(data);

            int statusCode = client.executeMethod(post);

            System.out.println(statusCode);    // must be 302

            if (statusCode > 400) {
                throw new EXZException("Invalid URL");
            }

            // handle response.
            String   cookieValue = null;
            Cookie[] cookies     = client.getState().getCookies();

            for (int i = 0; i < cookies.length; i++) {
                if (cookies[i].getName().equals(instanceName)) {
                    cookieValue = cookies[i].getValue();
                    System.out.println(cookieValue);

                    break;
                }
            }

            if (cookieValue == null) {
                throw new EXZException("Invalid username/password combination");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
*/
}
