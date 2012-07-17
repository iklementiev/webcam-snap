/******************************************************
                  DirectShow .NET
		      netmaster@swissonline.ch
			  
			  Modified by Dan Glass 03/02/03
*******************************************************/
 
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

using DShowNET;
using DShowNET.Device;


namespace webcamsnap
{

	/// <summary> Summary description for MainForm. </summary>
	public class Capture : ISampleGrabberCB
	{

		// so we can wait for the async job to finish
		private ManualResetEvent reset = new ManualResetEvent(false);

		// the image - needs to be global because it is fetched with an async method call
		private Image image;

		/// <summary> base filter of the actually used video devices. </summary>
		private IBaseFilter				capFilter;

		/// <summary> graph builder interface. </summary>
		private IGraphBuilder			graphBuilder;

		/// <summary> capture graph builder interface. </summary>
		private ICaptureGraphBuilder2	capGraph;
		private ISampleGrabber			sampGrabber;

		/// <summary> control interface. </summary>
		private IMediaControl			mediaCtrl;

		/// <summary> grabber filter interface. </summary>
		private IBaseFilter				baseGrabFlt;

		/// <summary> structure describing the bitmap to grab. </summary>
		private	VideoInfoHeader			videoInfoHeader;

		/// <summary> buffer for bitmap data. </summary>
		private	byte[]					savedArray;

		/// <summary> list of installed video devices. </summary>
		private ArrayList				capDevices;

		private Capture()
		{
		}

		public static Image GetImage()
		{
			Capture cam = new Capture();

			try
			{
				cam.Init();
				cam.CaptureImage();
				cam.CloseInterfaces();
			
				Image image = (Image)cam.image.Clone();
				cam = null;
				return image;
			}
			catch (Exception ex)
			{
				cam.CloseInterfaces();
				Debug.WriteLine(ex);
				return null;
			}
		}

		public static Image GetImage(DsDevice device)
		{
			Capture cam = new Capture();

			try
			{
				cam.Init(device);
				cam.CaptureImage();
				cam.CloseInterfaces();
			
				return cam.image;
			}
			catch (Exception ex)
			{
				cam.CloseInterfaces();
				Debug.WriteLine(ex);
				return null;
			}
		}

		public static DsDevice[] GetInterfaceList()
		{
			ArrayList devices = null;

			if( ! DsUtils.IsCorrectDirectXVersion() )
			{
				throw new Exception("DirectX 8.1 NOT installed!");
			}

			if( ! DsDev.GetDevicesOfCat( FilterCategory.VideoInputDevice, out devices ) )
			{
				//throw new Exception("No video capture devices found!");
				devices = new ArrayList();
			}

			return (DsDevice[])devices.ToArray(typeof(DsDevice));
		}

		private void CaptureImage()
		{
		
			int hr;
			if( sampGrabber == null )
				return;

			if( savedArray == null )
			{
				int size = videoInfoHeader.BmiHeader.ImageSize;
				// sanity check
				if( (size < 1000) || (size > 16000000) )
					return;
				savedArray = new byte[ size + 64000 ];
			}

			hr = sampGrabber.SetCallback( this, 1 );
			
			if ( ! reset.WaitOne(5000,false) )
				throw new Exception("Timeout waiting to get picture");
		}

		private void Init()
		{
			if( ! DsUtils.IsCorrectDirectXVersion() )
			{
				throw new Exception("DirectX 8.1 NOT installed!");
			}

			if( ! DsDev.GetDevicesOfCat( FilterCategory.VideoInputDevice, out capDevices ) )
			{
				throw new Exception("No video capture devices found!");
			}

			DsDevice dev = capDevices[0] as DsDevice;

			StartupVideo( dev.Mon );
		}

		private void Init(DsDevice device)
		{
			// store it for clean up.
			capDevices = new ArrayList();
			capDevices.Add(device);

			StartupVideo( device.Mon );
		}

		/// <summary> capture event, triggered by buffer callback. </summary>
		private void OnCaptureDone()
		{
			if( sampGrabber == null )
				return;

			int w = videoInfoHeader.BmiHeader.Width;
			int h = videoInfoHeader.BmiHeader.Height;
			if( ((w & 0x03) != 0) || (w < 32) || (w > 4096) || (h < 32) || (h > 4096) )
				return;

			int stride = w * 3;

			GCHandle handle = GCHandle.Alloc( savedArray, GCHandleType.Pinned );
			int scan0 = (int) handle.AddrOfPinnedObject();
			scan0 += (h - 1) * stride;
			image = new Bitmap( w, h, -stride, PixelFormat.Format24bppRgb, (IntPtr) scan0 );
			handle.Free();
			savedArray = null;
			reset.Set();
		}



		/// <summary> start all the interfaces, graphs and preview window. </summary>
		private bool StartupVideo( UCOMIMoniker mon )
		{
			int hr;
			if( ! CreateCaptureDevice( mon ) )
				return false;

			if( ! GetInterfaces() )
				return false;

			if( ! SetupGraph() )
				return false;
			
			hr = mediaCtrl.Run();
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			// will thow up input for tuner settings.
			//bool hasTuner = DsUtils.ShowTunerPinDialog( capGraph, capFilter, IntPtr.Zero );

			return true;
		}

		/// <summary> build the capture graph for grabber. </summary>
		private bool SetupGraph()
		{
			int hr;
			hr = capGraph.SetFiltergraph( graphBuilder );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			hr = graphBuilder.AddFilter( capFilter, "Ds.NET Video Capture Device" );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			
			// will thow up user input for quality
			//DsUtils.ShowCapPinDialog( capGraph, capFilter, IntPtr.Zero );

			AMMediaType media = new AMMediaType();
			media.majorType	= MediaType.Video;
			media.subType	= MediaSubType.RGB24;
			media.formatType = FormatType.VideoInfo;		// ???
			hr = sampGrabber.SetMediaType( media );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			hr = graphBuilder.AddFilter( baseGrabFlt, "Ds.NET Grabber" );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			Guid cat;
			Guid med;

			cat = PinCategory.Capture;
			med = MediaType.Video;
			hr = capGraph.RenderStream( ref cat, ref med, capFilter, null, baseGrabFlt ); // baseGrabFlt 

			media = new AMMediaType();
			hr = sampGrabber.GetConnectedMediaType( media );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );
			if( (media.formatType != FormatType.VideoInfo) || (media.formatPtr == IntPtr.Zero) )
				throw new NotSupportedException( "Unknown Grabber Media Format" );

			videoInfoHeader = (VideoInfoHeader) Marshal.PtrToStructure( media.formatPtr, typeof(VideoInfoHeader) );
			Marshal.FreeCoTaskMem( media.formatPtr ); media.formatPtr = IntPtr.Zero;

			hr = sampGrabber.SetBufferSamples( false );
			if( hr == 0 )
				hr = sampGrabber.SetOneShot( false );
			if( hr == 0 )
				hr = sampGrabber.SetCallback( null, 0 );
			if( hr < 0 )
				Marshal.ThrowExceptionForHR( hr );

			return true;
		}


		/// <summary> create the used COM components and get the interfaces. </summary>
		private bool GetInterfaces()
		{
			Type comType = null;
			object comObj = null;
			try 
			{
				comType = Type.GetTypeFromCLSID( Clsid.FilterGraph );
				if( comType == null )
					throw new NotImplementedException( @"DirectShow FilterGraph not installed/registered!" );
				comObj = Activator.CreateInstance( comType );
				graphBuilder = (IGraphBuilder) comObj; comObj = null;

				Guid clsid = Clsid.CaptureGraphBuilder2;
				Guid riid = typeof(ICaptureGraphBuilder2).GUID;
				comObj = DsBugWO.CreateDsInstance( ref clsid, ref riid );
				capGraph = (ICaptureGraphBuilder2) comObj; comObj = null;

				comType = Type.GetTypeFromCLSID( Clsid.SampleGrabber );
				if( comType == null )
					throw new NotImplementedException( @"DirectShow SampleGrabber not installed/registered!" );
				comObj = Activator.CreateInstance( comType );
				sampGrabber = (ISampleGrabber) comObj; comObj = null;

				mediaCtrl	= (IMediaControl)	graphBuilder;
				baseGrabFlt	= (IBaseFilter)		sampGrabber;
				return true;
			}
			catch( Exception ee )
			{
				if( comObj != null )
					Marshal.ReleaseComObject( comObj ); comObj = null;

				throw ee;
			}
		}

		/// <summary> create the user selected capture device. </summary>
		private bool CreateCaptureDevice( UCOMIMoniker mon )
		{
			object capObj = null;
			try 
			{
				Guid gbf = typeof( IBaseFilter ).GUID;
				mon.BindToObject( null, null, ref gbf, out capObj );
				capFilter = (IBaseFilter) capObj; capObj = null;
				return true;
			}
			catch( Exception ee )
			{
				if( capObj != null )
					Marshal.ReleaseComObject( capObj ); capObj = null;
				throw ee;
			}
		}



		/// <summary>
		/// MUST do this. 
		/// Notice alot of crap in here - I've had some problems over extending the bandwidth of
		/// the USB bas.
		/// </summary>
		private void CloseInterfaces()
		{
			int hr;
			try 
			{
				if( mediaCtrl != null )
				{
					hr = mediaCtrl.Stop();
					mediaCtrl = null;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}

			baseGrabFlt = null;

			try 
			{
				if( sampGrabber != null ) Marshal.ReleaseComObject( sampGrabber ); sampGrabber = null;
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}

			try 
			{
				if( capGraph != null ) Marshal.ReleaseComObject( capGraph ); capGraph = null;
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}

			try 
			{
				if( graphBuilder != null ) Marshal.ReleaseComObject( graphBuilder ); graphBuilder = null;
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}

			try 
			{
				if( capFilter != null ) Marshal.ReleaseComObject( capFilter ); capFilter = null;
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}

			try 
			{
				if( capDevices != null )
				{
					foreach( DsDevice d in capDevices )
						d.Dispose();

					capDevices = null;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
			}
		}

		/// <summary> sample callback, NOT USED. </summary>
		int ISampleGrabberCB.SampleCB( double SampleTime, IMediaSample pSample )
		{
			return 0;
		}

		/// <summary> buffer callback, COULD BE FROM FOREIGN THREAD. </summary>
		int ISampleGrabberCB.BufferCB( double SampleTime, IntPtr pBuffer, int BufferLen )
		{
			if( savedArray == null )
			{
				return 0;
			}

			if( (pBuffer != IntPtr.Zero) && (BufferLen > 1000) && (BufferLen <= savedArray.Length) )
				Marshal.Copy( pBuffer, savedArray, 0, BufferLen );

			this.OnCaptureDone();
			return 0;
		}
	}
}
