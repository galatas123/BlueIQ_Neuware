using System;
using System.Linq;
using System;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace BlueIQ_Neuware
{
    internal partial class About : Form
    {
        public About()
        {
            InitializeComponent();

            SetAssemblyInformation();
        }

        private void SetAssemblyInformation()
        {
            // Set Product Name
            var productNameAttribute = Assembly.GetExecutingAssembly()
                .GetCustomAttributes(typeof(AssemblyProductAttribute), false)
                .OfType<AssemblyProductAttribute>()
                .FirstOrDefault();

            labelProductName.Text = productNameAttribute != null ? productNameAttribute.Product : "BlueIQ Control";

            // Set Version
            Version? version = Assembly.GetExecutingAssembly().GetName().Version;
            labelVersion.Text = version != null ? $"Version: {version}" : "Version: 1.0";

            // Set Copyright
            var copyrightAttribute = Assembly.GetExecutingAssembly()
                .GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false)
                .OfType<AssemblyCopyrightAttribute>()
                .FirstOrDefault();

            labelCopyright.Text = copyrightAttribute != null ? copyrightAttribute.Copyright : "© 2023 Ingram Micro Services";

            // Set Company Name
            var companyAttribute = Assembly.GetExecutingAssembly()
                .GetCustomAttributes(typeof(AssemblyCompanyAttribute), false)
                .OfType<AssemblyCompanyAttribute>()
                .FirstOrDefault();

            labelCompanyName.Text = companyAttribute != null ? $"{companyAttribute.Company}" : "Programmer: Anil Chikmet Oglou";
        }
    }
}

