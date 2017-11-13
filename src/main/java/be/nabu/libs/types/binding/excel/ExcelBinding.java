package be.nabu.libs.types.binding.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

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
import be.nabu.utils.excel.ExcelUtils;
import be.nabu.utils.excel.FileType;

public class ExcelBinding implements MarshallableBinding {

	private boolean useHeader = true;
	
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
								Value<String> property = child.getProperty(AliasProperty.getInstance());
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
						row.add(((ComplexContent) item).get(child.getName()));
					}
					rows.add(row);
				}
				ExcelUtils.write(output, rows, "export", FileType.XLSX, null);
			}
		}
	}

	public boolean isUseHeader() {
		return useHeader;
	}

	public void setUseHeader(boolean useHeader) {
		this.useHeader = useHeader;
	}

}
