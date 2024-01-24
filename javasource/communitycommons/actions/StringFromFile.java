// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package communitycommons.actions;

import com.mendix.systemwideinterfaces.core.IMendixObject;
import communitycommons.StringUtils;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;

/**
 * Reads the contents form the provided file document, using the specified encoding, and returns it as string.
 */
public class StringFromFile extends CustomJavaAction<java.lang.String>
{
	/** @deprecated use source.getMendixObject() instead. */
	@java.lang.Deprecated(forRemoval = true)
	private final IMendixObject __source;
	private final system.proxies.FileDocument source;
	private final communitycommons.proxies.StandardEncodings encoding;

	public StringFromFile(
		IContext context,
		IMendixObject _source,
		java.lang.String _encoding
	)
	{
		super(context);
		this.__source = _source;
		this.source = _source == null ? null : system.proxies.FileDocument.initialize(getContext(), _source);
		this.encoding = _encoding == null ? null : communitycommons.proxies.StandardEncodings.valueOf(_encoding);
	}

	@java.lang.Override
	public java.lang.String executeAction() throws Exception
	{
		// BEGIN USER CODE
		Charset charset = StandardCharsets.UTF_8;
		if (this.encoding != null)
			charset = Charset.forName(this.encoding.name().replace('_', '-'));
		return StringUtils.stringFromFile(getContext(), source, charset);
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "StringFromFile";
	}

	// BEGIN EXTRA CODE
	// END EXTRA CODE
}
