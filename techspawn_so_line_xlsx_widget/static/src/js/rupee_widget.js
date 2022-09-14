odoo.define("rupee_widget", function(require) {
  "use strict";
  console.log("RUPEEEE")
  var field_registry = require("web.field_registry");
  var basic_fields = require("web.basic_fields");

  var RupeeWidgetFloat = basic_fields.FieldFloat.extend({
    template: "RupeeWidget",
    supportedFieldTypes: ['integer', 'float'],

    _render: function() {
      var intValue = parseFloat(this.value);
      var parseFloatValue = isNaN(intValue) ? 0 : intValue;
        var $input = this.$el.find("input");
        $input.val(intValue);
        this.$input = $input;
        this.$(".oe_field_rupee_float").text("₹ "+ intValue);
    }
  });

    var RupeeWidgetInteger = basic_fields.FieldInteger.extend({
    template: "RupeeWidgetInteger",

    _render: function() {
      var intValue = parseInt(this.value);
      var parseIntValue = isNaN(intValue) ? 0 : intValue;
        var $input = this.$el.find("input");
        $input.val(intValue);
        this.$input = $input;
        this.$(".oe_field_rupee_integer").text("₹ "+ intValue);
    }
  });

  field_registry.add("rupee_float", RupeeWidgetFloat);
  field_registry.add("rupee_integer", RupeeWidgetInteger);
});



