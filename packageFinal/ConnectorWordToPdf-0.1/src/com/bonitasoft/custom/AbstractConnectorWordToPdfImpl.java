package com.bonitasoft.custom;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorValidationException;

public abstract class AbstractConnectorWordToPdfImpl extends AbstractConnector {

	protected final static String FROM_INPUT_PARAMETER = "from";
	protected final static String TO_INPUT_PARAMETER = "to";
	protected final static String TMP_INPUT_PARAMETER = "tmp";
	protected final static String FILENAME_INPUT_PARAMETER = "fileName";
	protected final static String FILENAMEFINAL_INPUT_PARAMETER = "fileNameFinal";
	protected final static String MAPPING_INPUT_PARAMETER = "mapping";
	protected final Boolean OUT_OUTPUT_PARAMETER = "out";

	protected final java.lang.String getFrom() {
		return (java.lang.String) getInputParameter(FROM_INPUT_PARAMETER);
	}

	protected final java.lang.String getTo() {
		return (java.lang.String) getInputParameter(TO_INPUT_PARAMETER);
	}

	protected final java.lang.String getTmp() { return (java.lang.String) getInputParameter(TMP_INPUT_PARAMETER); }

	protected final java.lang.String getFileName() { return (java.lang.String) getInputParameter(FILENAME_INPUT_PARAMETER); }

	protected final java.lang.String getFileNameFinal() { return (java.lang.String) getInputParameter(FILENAMEFINAL_INPUT_PARAMETER); }

	protected final java.util.List getMapping() { return (java.util.List) getInputParameter(MAPPING_INPUT_PARAMETER); }

	protected final void setOut(java.lang.Boolean out) {
		setOutputParameter(OUT_OUTPUT_PARAMETER, out);
	}

	@Override
	public void validateInputParameters() throws ConnectorValidationException {
		try {
			getFrom();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("from type is invalid");
		}
		try {
			getTo();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("to type is invalid");
		}
		try {
			getTmp();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("tmp type is invalid");
		}
		try {
			getFileName();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("fileName type is invalid");
		}
		try {
			getFileNameFinal();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("fileNameFinal type is invalid");
		}
		try {
			getMapping();
		} catch (ClassCastException cce) {
			throw new ConnectorValidationException("mapping type is invalid");
		}
	}
}
