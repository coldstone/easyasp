##EasyASP 更新日志
- [2014/05/12] 新增：表单验证函数，采用链式操作，可灵活验证各种类型数据并支持设置默认值、自定义错误提示信息、弹出错误信息对话框。内置超过30种验证规则，并支持表单一致性验证和验证码验证。
- [2014/05/09] 修正：Easp.Str.HtmlFormat 调用出错的bug。
- [2014/05/08] 新增：Easp.Str.Format 支持超级变量静态标签的引用替换。如 Easp.Str.Format("this is a {=name} test for {0} and {1}", arr) 会自动将 {=name} 的值替换为 Easp.Var("name") 的值。
- [2014/05/08] 新增：超级变量支持全局无限级嵌套静态标签。如设置 Easp.Var("type") = "typename is {=name}" 时，就可直接引用 Easp.Var("name") 的值，而在 Easp.Var("name") 中还可以引用别的超级变量，支持无限级引用。