import logging
from typing import Any, Dict, Optional, Union

logger = logging.getLogger("InventorAutomation.Properties")


class PropertyManager:
    """Manages Autodesk Inventor iProperties (standard and user-defined custom metadata)."""

    SUMMARY_SET = "Inventor Summary Information"
    DOC_SUMMARY_SET = "Inventor Document Summary Information"
    DESIGN_TRACKING_SET = "Design Tracking Properties"
    USER_DEFINED_SET = "Inventor User Defined Properties"

    def __init__(self, doc: Any):
        self.doc = doc
        try:
            self.property_sets = doc.PropertySets
        except AttributeError as e:
            logger.error(f"Provided document does not have PropertySets: {e}")
            raise ValueError("Document must expose PropertySets.") from e

    def get_property(self, set_name: str, prop_name: str, default: Any = None) -> Any:
        """Retrieves the value of a property from a specific property set."""
        try:
            prop_set = self.property_sets.Item(set_name)
            prop = prop_set.Item(prop_name)
            return prop.Value
        except Exception:
            return default

    def set_property(self, set_name: str, prop_name: str, value: Any) -> bool:
        """Sets the value of an existing standard property."""
        try:
            prop_set = self.property_sets.Item(set_name)
            prop = prop_set.Item(prop_name)
            prop.Value = value
            logger.debug(f"Set property [{set_name}].'{prop_name}' = '{value}'")
            return True
        except Exception as e:
            logger.warning(f"Failed to set standard property [{set_name}].'{prop_name}': {e}")
            return False

    def get_custom_property(self, prop_name: str, default: Any = None) -> Any:
        """Retrieves a user-defined custom property value."""
        return self.get_property(self.USER_DEFINED_SET, prop_name, default)

    def set_custom_property(self, prop_name: str, value: Any) -> bool:
        """Sets or creates a user-defined custom iProperty."""
        try:
            custom_set = self.property_sets.Item(self.USER_DEFINED_SET)
            try:
                prop = custom_set.Item(prop_name)
                prop.Value = value
            except Exception:
                # Add new custom property
                custom_set.Add(value, prop_name)
            logger.debug(f"Set custom property '{prop_name}' = '{value}'")
            return True
        except Exception as e:
            logger.warning(f"Failed to set custom property '{prop_name}': {e}")
            return False

    def set_standard_properties(
        self,
        part_number: Optional[str] = None,
        description: Optional[str] = None,
        author: Optional[str] = None,
        designer: Optional[str] = None,
        company: Optional[str] = None,
        revision: Optional[str] = None,
        vendor: Optional[str] = None,
        project: Optional[str] = None,
    ) -> None:
        """Configures all common engineering and title block properties in one call."""
        if part_number is not None:
            self.set_property(self.DESIGN_TRACKING_SET, "Part Number", str(part_number))
        if description is not None:
            self.set_property(self.DESIGN_TRACKING_SET, "Description", str(description))
        if author is not None:
            self.set_property(self.SUMMARY_SET, "Author", str(author))
        if designer is not None:
            self.set_property(self.DESIGN_TRACKING_SET, "Designer", str(designer))
        if company is not None:
            self.set_property(self.DOC_SUMMARY_SET, "Company", str(company))
        if revision is not None:
            self.set_property(self.SUMMARY_SET, "Revision Number", str(revision))
        if vendor is not None:
            self.set_property(self.DESIGN_TRACKING_SET, "Vendor", str(vendor))
        if project is not None:
            self.set_property(self.DESIGN_TRACKING_SET, "Project", str(project))

    def get_all_properties(self) -> Dict[str, Dict[str, Any]]:
        """Returns a nested dictionary of all properties across all property sets."""
        result: Dict[str, Dict[str, Any]] = {}
        try:
            for i in range(1, self.property_sets.Count + 1):
                prop_set = self.property_sets.Item(i)
                set_dict: Dict[str, Any] = {}
                for j in range(1, prop_set.Count + 1):
                    try:
                        p = prop_set.Item(j)
                        set_dict[p.Name] = p.Value
                    except Exception:
                        continue
                result[prop_set.Name] = set_dict
        except Exception as e:
            logger.warning(f"Error enumerating property sets: {e}")
        return result
