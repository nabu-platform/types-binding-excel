/*
* Copyright (C) 2017 Alexander Verbruggen
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU Lesser General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
* GNU Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public License
* along with this program. If not, see <https://www.gnu.org/licenses/>.
*/

package be.nabu.libs.types.binding.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.TimeZone;

import be.nabu.libs.property.api.Value;
import be.nabu.libs.types.CollectionHandlerFactory;
import be.nabu.libs.types.ComplexContentWrapperFactory;
import be.nabu.libs.types.TypeUtils;
import be.nabu.libs.types.api.CollectionHandlerProvider;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.api.ComplexType;
import be.nabu.libs.types.api.Element;
import be.nabu.libs.types.api.SimpleType;
import be.nabu.libs.types.binding.api.MarshallableBinding;
import be.nabu.libs.types.java.BeanType;
import be.nabu.libs.types.properties.AliasProperty;
import be.nabu.libs.types.properties.LabelProperty;
import be.nabu.utils.excel.ExcelUtils;
import be.nabu.utils.excel.FileType;

public class ExcelBinding implements MarshallableBinding {

	private boolean useHeader = true;
	private boolean allowFormulas = false;
	private TimeZone timezone;
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public void marshal(OutputStream output, ComplexContent content, Value<?>... values) throws IOException {
		for (Element<?> element : TypeUtils.getAllChildren(content.getType())) {
			if (element.getType() instanceof ComplexType && element.getType().isList(element.getProperties())) {
				Object object = content.get(element.getName());
				// no data, skip
				if (object == null) {
					continue;
				}
				CollectionHandlerProvider handler = CollectionHandlerFactory.getInstance().getHandler().getHandler(object.getClass());

				List<Object> rows = new ArrayList<Object>();
				Collection<Element<?>> children = null;
				boolean firstItem = true;
				for (Object item : handler.getAsIterable(object)) {
					if (!(item instanceof ComplexContent)) {
						item = ComplexContentWrapperFactory.getInstance().getWrapper().wrap(item);
					}
					if (firstItem) {
						ComplexType recordType = (ComplexType) element.getType();
						// if we are dealing with objects, use the runtime definition
						if (recordType instanceof BeanType && ((BeanType) recordType).getBeanClass().equals(Object.class)) {
							recordType = ((ComplexContent) item).getType();
						}
						children = TypeUtils.getAllChildren(recordType);

						if (useHeader) {
							// write header hashtag
							List<Object> header = new ArrayList<Object>();
							for (Element<?> child : children) {
								if (!(child.getType() instanceof SimpleType)) {
									continue;
								}
								Value<String> property = child.getProperty(LabelProperty.getInstance());
								if (property == null) {
									property = child.getProperty(AliasProperty.getInstance());
								}
								header.add(property == null ? child.getName() : property.getValue());
							}
							rows.add(header);
						}
						firstItem = false;
					}
					
					List<Object> row = new ArrayList<Object>();
					for (Element<?> child : children) {
						if (!(child.getType() instanceof SimpleType)) {
							continue;
						}
						Object e = ((ComplexContent) item).get(child.getName());
						if (e instanceof String && !allowFormulas && e.toString().startsWith("=")) {
							// escape it
							e = "'" + e;
						}
//						// if we have a string but it contains only numbers, we want to explicitly escape it
//						else if (String.class.isAssignableFrom(((SimpleType) child.getType()).getInstanceClass()) && e instanceof String && ((String) e).matches("^[0-9.]+$")) {
//							e = "'" + e;
//						}
						row.add(e);
					}
					rows.add(row);
				}
				Value<String> property = element.getProperty(LabelProperty.getInstance());
				if (property == null) {
					property = element.getProperty(AliasProperty.getInstance());
				}
				ExcelUtils.write(output, rows, property == null ? element.getName() : property.getValue(), FileType.XLSX, null, timezone);
			}
		}
	}

	public boolean isUseHeader() {
		return useHeader;
	}

	public void setUseHeader(boolean useHeader) {
		this.useHeader = useHeader;
	}

	public TimeZone getTimezone() {
		return timezone;
	}

	public void setTimezone(TimeZone timezone) {
		this.timezone = timezone;
	}
}
