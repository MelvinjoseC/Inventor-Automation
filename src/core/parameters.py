import logging
from typing import Any, Dict, Optional, Union

logger = logging.getLogger("InventorAutomation.Parameters")


class ParameterManager:
    """Manages User Parameters and Model Parameters in Autodesk Inventor documents.
    
    Provides programmatic read/write access to parametric CAD dimensions and equations.
    """

    def __init__(self, doc: Any):
        self.doc = doc
        try:
            self.cd = doc.ComponentDefinition
            self.parameters = self.cd.Parameters
            self.user_parameters = self.parameters.UserParameters
            self.model_parameters = self.parameters.ModelParameters
        except AttributeError as e:
            logger.error(f"Provided document does not expose ComponentDefinition.Parameters: {e}")
            raise ValueError("Document must be a Part or Assembly document.") from e

    def set_user_parameter(
        self,
        name: str,
        expression: Union[str, float, int],
        units: str = "mm",
    ) -> Any:
        """Creates or updates a user parameter with a given name, expression, and unit."""
        expr_str = f"{expression} {units}".strip() if isinstance(expression, (int, float)) else str(expression)
        
        try:
            # Check if parameter already exists
            param = self.user_parameters.Item(name)
            logger.debug(f"Updating existing parameter '{name}' to '{expr_str}'")
            param.Expression = expr_str
            return param
        except Exception:
            # Add new parameter
            try:
                logger.debug(f"Adding new user parameter '{name}' with expression '{expr_str}'")
                param = self.user_parameters.AddByExpression(name, expr_str, units)
                return param
            except Exception as ex:
                logger.error(f"Failed to add/update user parameter '{name}': {ex}")
                raise

    def get_parameter(self, name: str) -> Optional[Any]:
        """Retrieves a parameter by name (searches User and Model parameters)."""
        try:
            return self.parameters.Item(name)
        except Exception:
            return None

    def get_parameter_value(self, name: str) -> Optional[float]:
        """Gets evaluated numeric value in internal database units (e.g. cm for length)."""
        param = self.get_parameter(name)
        return param.Value if param else None

    def get_parameter_expression(self, name: str) -> Optional[str]:
        """Gets user-facing parameter expression string (e.g. '50.0 mm')."""
        param = self.get_parameter(name)
        return param.Expression if param else None

    def update_parameters(self, params: Dict[str, Union[str, float]], default_units: str = "mm") -> None:
        """Batch updates or creates multiple parameters and triggers model update."""
        for name, expr in params.items():
            self.set_user_parameter(name, expr, units=default_units)
        
        # Trigger document update to propagate parameter changes across features
        try:
            self.doc.Update2(True)
            logger.info(f"Updated {len(params)} parameters and triggered document regeneration.")
        except Exception as e:
            logger.warning(f"Could not trigger automatic document regeneration: {e}")

    def get_all_user_parameters(self) -> Dict[str, Dict[str, Any]]:
        """Returns a dictionary summary of all user parameters in the document."""
        result: Dict[str, Dict[str, Any]] = {}
        try:
            count = self.user_parameters.Count
            for i in range(1, count + 1):
                p = self.user_parameters.Item(i)
                result[p.Name] = {
                    "expression": p.Expression,
                    "value": p.Value,
                    "units": getattr(p, "Units", ""),
                    "comment": getattr(p, "Comment", ""),
                }
        except Exception as e:
            logger.warning(f"Error enumerating user parameters: {e}")
        return result
