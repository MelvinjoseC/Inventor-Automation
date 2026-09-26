import pytest
from unittest.mock import MagicMock
from src.core.parameters import ParameterManager
from src.core.properties import PropertyManager


def test_parameter_manager():
    doc = MagicMock()
    cd = doc.ComponentDefinition
    params = cd.Parameters
    user_params = params.UserParameters
    
    # Mock existing parameter lookup and addition
    mock_param = MagicMock()
    mock_param.Name = "Length"
    mock_param.Expression = "100 mm"
    mock_param.Value = 10.0
    
    # Return mock_param when requested
    user_params.Item.return_value = mock_param
    user_params.Count = 1
    user_params.Item.side_effect = lambda idx_or_name: mock_param if idx_or_name in ("Length", 1) else (_ for _ in ()).throw(KeyError)
    
    mgr = ParameterManager(doc)
    
    # Update existing parameter
    mgr.set_user_parameter("Length", "120 mm")
    assert mock_param.Expression == "120 mm"
    
    # Add new parameter (simulate lookup fail then AddByExpression success)
    def item_lookup(k):
        if k == "Length":
            return mock_param
        raise KeyError(k)
        
    user_params.Item.side_effect = item_lookup
    new_param = MagicMock()
    user_params.AddByExpression.return_value = new_param
    res = mgr.set_user_parameter("Width", 50, "mm")
    user_params.AddByExpression.assert_called_with("Width", "50 mm", "mm")
    assert res == new_param
    
    # Batch update
    mgr.update_parameters({"Length": "150 mm", "Width": "60 mm"})
    doc.Update2.assert_called_with(True)


def test_property_manager():
    doc = MagicMock()
    prop_sets = doc.PropertySets
    
    summary_set = MagicMock()
    summary_set.Name = "Inventor Summary Information"
    design_set = MagicMock()
    design_set.Name = "Design Tracking Properties"
    custom_set = MagicMock()
    custom_set.Name = "Inventor User Defined Properties"
    
    sets_dict = {
        "Inventor Summary Information": summary_set,
        "Design Tracking Properties": design_set,
        "Inventor User Defined Properties": custom_set,
    }
    prop_sets.Item.side_effect = lambda name: sets_dict[name]
    
    part_no_prop = MagicMock()
    part_no_prop.Value = "PART-001"
    design_set.Item.return_value = part_no_prop
    
    author_prop = MagicMock()
    author_prop.Value = "Senior CAD Engineer"
    summary_set.Item.return_value = author_prop
    
    mgr = PropertyManager(doc)
    
    # Set standard properties
    mgr.set_standard_properties(part_number="BM-999", description="Engine Bracket", author="CAD Team")
    design_set.Item.assert_any_call("Part Number")
    design_set.Item.assert_any_call("Description")
    summary_set.Item.assert_any_call("Author")
    
    # Custom property addition
    custom_set.Item.side_effect = KeyError("Not found")
    mgr.set_custom_property("QA_Status", "Approved")
    custom_set.Add.assert_called_with("Approved", "QA_Status")
