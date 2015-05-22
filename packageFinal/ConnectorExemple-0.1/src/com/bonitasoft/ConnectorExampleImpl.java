/**
 *
 */
package com.bonitasoft;

import com.bonitasoft.ConnectorLib;
import java.util.logging.Logger;
import org.bonitasoft.engine.connector.ConnectorException;

/**
 *The connector execution will follow the steps
 * 1 - setInputParameters() --> the connector receives input parameters values
 * 2 - validateInputParameters() --> the connector can validate input parameters values
 * 3 - connect() --> the connector can establish a connection to a remote server (if necessary)
 * 4 - executeBusinessLogic() --> execute the connector
 * 5 - getOutputParameters() --> output are retrieved from connector
 * 6 - disconnect() --> the connector can close connection to remote server (if any)
 */
public class ConnectorExampleImpl extends AbstractConnectorExampleImpl {

    @Override
    protected void executeBusinessLogic() throws ConnectorException{
        //Get access to the connector input parameters
        //getWho();

        //var
        Boolean trace = true;
        String headerLog = "[Log : "+this.getClass().getName()+"]";

        //Init logger
        Logger logger = Logger.getLogger("com.bonitasoft.groovy");
        logger.info(headerLog + "Execute method executeBusinessLogic.");

        if(trace){
            logger.info(headerLog + "who : " + getWho());
        }

        String myReturn = ConnectorLib.sayHello(getWho());

        setOut(myReturn);
    }

    @Override
    public void connect() throws ConnectorException{
        //[Optional] Open a connection to remote server

    }

    @Override
    public void disconnect() throws ConnectorException{
        //[Optional] Close connection to remote server

    }

}
