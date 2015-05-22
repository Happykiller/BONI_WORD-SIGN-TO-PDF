package com.bonitasoft;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorValidationException;

public abstract class AbstractConnectorExampleImpl extends AbstractConnector {

	protected final static String WHO_INPUT_PARAMETER = "who";
	protected final String OUT_OUTPUT_PARAMETER = "out";

	protected final java.lang.String getWho() {
		return (java.lang.String) getInputParameter(WHO_INPUT_PARAMETER);
	}

	protected final void setOut(java.lang.String out) {
		setOutputParameter(OUT_OUTPUT_PARAMETER, out);
	}

	@Override
	public void validateInputParameters() throws ConnectorValidationException {
		try {
			getWho();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("who type is invalid");
		}
	}
}
