public class TestController : Controller
{
    // GET: Test
    public ActionResult Index()
    {
        return View();
    }

    // GET: Test/Process
    public ActionResult Process()
    {
        return View(new TestObjectRequest());
    }

    // POST: Certigy/Create
    [HttpPost]
    public ActionResult Process(TestObjectRequest request)
    {
        try
        {
            var testObject = ComClassWrapper.ComObject("TestObjApplication.TestObject");
            var requestWrapper = new System.Runtime.InteropServices.VariantWrapper(request.ToString().ToUpper());
            int errorCode = testObject.getTestWebData(ref requestWrapper);
            ViewBag.TestResponseCode = errorCode;
            ViewBag.TestResponse = requestWrapper.WrappedObject.ToString();
            request.LoadFromString(requestWrapper.WrappedObject.ToString());
            return View(request);
        }
        catch
        {
            throw;
            //return View();
        }
    }
}
