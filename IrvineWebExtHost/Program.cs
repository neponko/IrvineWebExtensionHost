using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace IrvineWebExtHost
{
	class Program
	{
		[DataContract]
		class InputData
		{
			[DataMember]
			public string	action		= null;
			[DataMember]
			public string	url			= null;
			[DataMember]
			public string	filename	= null;
			[DataMember]
			public string	referer		= null;
			[DataMember]
			public object	flag		= false;

			[DataMember(Name="object")]
			public string	comObject	= null;
			[DataMember]
			public string	method		= null;
			[DataMember]
			public List<object>	args	= null;
		}

		[DataContract]
		class OutputData
		{
			[DataMember]
			public double	version		= 0.1;		// version
			[DataMember]
			public string	status		= null;
			[DataMember]
			public string	message		= null;
		}



		static void Main(string[] args)
		{
			Console.OutputEncoding = new UTF8Encoding();

			InputData inData;
			try {
				using (var stdin = new BinaryReader(Console.OpenStandardInput())) {
					var inLength = stdin.ReadInt32();
					var bin = stdin.ReadBytes(inLength);

					var serializer = new DataContractJsonSerializer(typeof(InputData));
					using (var memstream = new MemoryStream(bin)) {
						inData = serializer.ReadObject(memstream) as InputData;
					}
				}
			} catch (Exception e) {
				Console.Error.WriteLine("Failed to read json from stdin.\n{0}", e.Message);
				return;
			}

			bool fail = false;
			if (inData.action == "Version") {
				ResultJson(fail);
				return;
			}

			if (string.IsNullOrWhiteSpace(inData.comObject))
				inData.comObject = "Irvine.Api";

			Type typeIrvine;
			dynamic irvine = CreateActiveXObject(inData.comObject, out typeIrvine);
			if (irvine == null) {
				Console.Error.WriteLine("Unknown ActiveX Object: \"{0}\"", inData.comObject);
				fail = true;
				ResultJson(fail);
				return;
			}

			try {
				switch (inData.action) {
				case "Download":
					// flag: [0:通常 1:ダイアログ 2:フォルダダイアログ 3:すぐにダウンロード 4:すぐにダイアログ 5:すぐにキューフォルダ]
					irvine.Download(inData.url, inData.flag);
					break;

				case "AddUrlAndReferer":
					NormalizeCurrentFolder(irvine);
					// flag: 0=全登録 1=選択ダイアログ
					irvine.AddUrlAndReferer(inData.url, inData.referer, inData.flag);
					break;

				case "CreateQueueItem":
					NormalizeCurrentFolder(irvine);
					// flag: 確認するかどうか
					irvine.CreateQueueItem(inData.url, inData.flag);
					break;

				case "ImportLinks":
					NormalizeCurrentFolder(irvine);
					// flag: 確認するかどうか
					irvine.ImportLinks(inData.url, inData.flag);
					break;

				case "AddItem":
					{
						Type typeItem;
						dynamic irvineItem = CreateActiveXObject("Irvine.Item", out typeItem);
						if (irvineItem == null) {
							fail = true;
							break;
						}

						irvineItem.Url = inData.url;
						if (!string.IsNullOrWhiteSpace(inData.filename))
							irvineItem.Filename = inData.filename;
						if (!string.IsNullOrWhiteSpace(inData.referer))
							irvineItem.Referer = inData.referer;

						NormalizeCurrentFolder(irvine);
						irvine.Current.Additem(irvineItem);
						Marshal.ReleaseComObject(irvineItem);
					}
					break;

				case null:
				case "":
					// { "object":"Irvine.Api", "method":"Download", "args":["http://...", 1] } 
					typeIrvine.InvokeMember(inData.method, BindingFlags.InvokeMethod, null, irvine, inData.args.ToArray());
					break;

				default:
					Console.Error.WriteLine("Unimplement action: \"{0}\"", inData.action);
					fail = true;
					break;
				}

			} catch (Exception e) {
				Console.Error.WriteLine("COM call failed.\n{0}", e.Message);
				fail = true;
			}

			Marshal.ReleaseComObject(irvine);
			ResultJson(fail);
		}


		static dynamic CreateActiveXObject(string id, out Type retType)
		{
			Type t;
			dynamic obj = null;

			try {
				t = Type.GetTypeFromProgID(id);
				if (t != null)
					obj = Activator.CreateInstance(t);
			} catch {
				t = null;
			}

			if (obj == null)
				Console.Error.WriteLine("Can't connect to Irvine.");

			retType = t;
			return obj;
		}


		static void NormalizeCurrentFolder(dynamic irvine)
		{
			if (irvine.Current.IsTrash())
				irvine.CurrentQueueFolder = "/Default";
		}


		static void ResultJson(bool fail = false, string message = null)
		{
			var outData = new OutputData() {
				status = fail ? "fail" : "success",
				message = message == null ? string.Empty : message
			};

			try {
				var serializer = new DataContractJsonSerializer(typeof(OutputData));
				using (var memstream = new MemoryStream()) {
					serializer.WriteObject(memstream, outData);
					using (var stdout = new BinaryWriter(Console.OpenStandardOutput())) {
						stdout.Write((Int32)memstream.Length);
						memstream.WriteTo(stdout.BaseStream);
					}
				}
			} catch {
			}
		}
	}
}
