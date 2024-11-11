using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace MicromanagerMcGhee {
    public partial class App : Application {

        protected override void OnStartup(StartupEventArgs e) {

            // Software Rendering
            base.OnStartup(e);
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            // Virtual Desktop

        }
    }
}
