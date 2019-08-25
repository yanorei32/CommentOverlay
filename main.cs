using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

static class CommentServer {
	static List<WebSocket> wsClients = new List<WebSocket>();

	static void logging(string module, string log) {
		Console.WriteLine("[{0, -8}] {1}", module, log);
	}

	class Comment {
		public string Text { get; private set; }
		public int DistFromRight { get; private set; }
		public int LinePos { get; private set; }
		public int Width { get; set; }

		public void Move(int dist) {
			DistFromRight += dist;
		}

		public Comment(string text, int linePos) {
			Text			= text;
			LinePos			= linePos;
			Width			= -1;
			DistFromRight	= 0;
		}
	}

	class Form1 : Form {
		public List<Comment> comments = new List<Comment>();
		const int VISIBLE_SEC = 3;
		const int TOP_OFFSET = 30;
		const int MARGIN_PER_LINE = 10;
		const int MAX_COMMENT_WIDTH = (int)(3840 * 1.5);
		const int WARN_FRAMES = 5;

		string CENTER_STRING;
		int CENTER_STRING_WIDTH, LINE_HEIGHT, FPS;
		bool firstFrame = true;
		Font FONT;

		void keyDown(object sender, KeyEventArgs e) {
			PowerPoint.Application p;

			try {
				p = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
			} catch { 
				return;
			}

			if (p == null) return;

			if (e.KeyCode == Keys.Q) {
				if (0 < p.SlideShowWindows.Count) {
					p.SlideShowWindows[1].View.Exit();
				}

				this.Close();
				return;
			}

			if (p.SlideShowWindows.Count < 1) {
				p.ActivePresentation.SlideShowSettings.Run();
				return;
			}

			switch (e.KeyCode) {
				case Keys.N:
					p.SlideShowWindows[1].View.Next();
					break;

				case Keys.P:
					p.SlideShowWindows[1].View.Previous();
					break;
			}
		}

		void drawComments(Graphics g) {
			firstFrame = false;

			g.TextRenderingHint = TextRenderingHint.SingleBitPerPixel;

			g.DrawString(
				CENTER_STRING,
				FONT,
				Brushes.Black,
				Size.Width / 2 - CENTER_STRING_WIDTH / 2,
				Size.Height - LINE_HEIGHT
			);

			lock(comments) {
				var deletionQueue = new List<Comment>();

				foreach (var c in comments) {
					if (c.Width == -1) {
						c.Width = TextRenderer.MeasureText(c.Text, FONT).Width;
					}

					if (c.Width >= MAX_COMMENT_WIDTH) {
						deletionQueue.Add(c);
						logging("Renderer", "MAX_COMMENT_WIDTH LIMIT!");
						continue;
					}

					if (Size.Width + c.Width - c.DistFromRight < 0) {
						deletionQueue.Add(c);
						continue;
					}

					var posV = TOP_OFFSET + (LINE_HEIGHT + MARGIN_PER_LINE) * c.LinePos;

					if (Size.Height <= posV) {
						deletionQueue.Add(c);
						logging("Renderer", "VERTICAL LIMIT!");
						continue;
					}

					g.DrawString(
						c.Text, FONT,
						Brushes.MediumSeaGreen,
						Size.Width - c.DistFromRight, posV
					);

					c.Move( (Size.Width + c.Width) / (FPS * VISIBLE_SEC) );
				}


				foreach (var c in deletionQueue) comments.Remove(c);
			}

		}

		void InitializeComponent() {
			var t = new Timer();
			t.Interval = 1000 / FPS;
			t.Tick += (sender, e) => {
				if (comments.Count == 0 && !firstFrame) return;
				Invalidate();
			};

			foreach (var s in Screen.AllScreens) {
				Location	= s.Bounds.Location;
				Size		= s.Bounds.Size;
				if (!s.Primary) break;
			}

			TopMost			= true;
			StartPosition	= FormStartPosition.Manual;
			WindowState		= FormWindowState.Maximized;
			TransparencyKey	= BackColor;
			FormBorderStyle	= FormBorderStyle.None;
			DoubleBuffered	= true;

			KeyDown		+= keyDown;
			Paint		+= (sender, e) => {
				drawComments(e.Graphics);
			};
			FormClosing += (sender, e) => {
				Parallel.ForEach(wsClients, async ws => {
					if (ws.State == WebSocketState.Open) {
						await ws.CloseAsync(
							WebSocketCloseStatus.NormalClosure,
							"Server stopped",
							System.Threading.CancellationToken.None
						);
					}
				});
			};

			t.Start();
		}

		public Form1(string centerString, int fontSize, int fps) {
			FONT			= new Font("Myrica M", fontSize, FontStyle.Bold);
			CENTER_STRING	= centerString;
			FPS				= fps;

			Size s = TextRenderer.MeasureText(centerString, FONT);
			LINE_HEIGHT			= s.Height;
			CENTER_STRING_WIDTH	= s.Width;

			InitializeComponent();
			StartServer(this);
		}

	}

	static async void StartServer(Form1 f) {
		string ip;
		var l = new HttpListener();
		l.Prefixes.Add("http://+:6928/");
		l.Start();
		for (;;) {
			var lctx = await l.GetContextAsync();
			ip = lctx.Request.RemoteEndPoint.Address.ToString();

			if (lctx.Request.IsWebSocketRequest) {
				ProcessRequest(f, lctx, ip);
				continue;
			}

			var r = lctx.Response;
			r.StatusCode = 200;

			var b = File.ReadAllBytes("index.html");
			r.OutputStream.Write(b, 0, b.Length);

			r.Close();
			logging("Web", "get html: " + ip);
		}
	}

	static async void ProcessRequest(Form1 f, HttpListenerContext hlc, string ip) {
		var ws = (await hlc.AcceptWebSocketAsync(subProtocol:null)).WebSocket;

		var closing = false;
		wsClients.Add(ws);
		logging("Web", "ws session create: " + ip);
		while (ws.State == WebSocketState.Open) {
			try {
				var buf = new byte[1024];

				var ret = await ws.ReceiveAsync(
					new ArraySegment<byte>(buf),
					System.Threading.CancellationToken.None
				);

				if (!ret.EndOfMessage) {
					logging("Web", "Too big message recv: " + ip);
					closing = true;
					await ws.CloseAsync(
						WebSocketCloseStatus.MessageTooBig,
						"Too big message recv",
						System.Threading.CancellationToken.None
					);
					continue;
				}

				if (ret.MessageType == WebSocketMessageType.Close) {
					logging("Web", "ws session closed: " + ip);
					break;
				}

				if (ret.MessageType != WebSocketMessageType.Text) continue;

				var comment = Encoding.UTF8.GetString(buf).TrimEnd('\0');

				logging("Web", string.Format("recv ({0}): {1}", ip, comment));

				lock(f.comments) {
					var pos = 0;
					while(f.comments.Exists(c => pos == c.LinePos)) pos++;
					f.comments.Add(new Comment(comment, pos));
				}
			} catch (Exception e) {
				if (closing) continue;
				logging("Web", "ws session abort: " + ip + e.ToString());
				break;
			}
		}

		wsClients.Remove(ws);
		ws.Dispose();
		logging("Web", "ws session removed: " + ip);
	}

	static void Main(string[] Args) {
		var fps = Args.Length >= 3 ? Int32.Parse(Args[2]) : 24;
		var fontSize = Args.Length >= 2 ? Int32.Parse(Args[1]) : 56;
		var centerStr = Args.Length >= 1 ? Args[0] : " ";

		Application.EnableVisualStyles();
		Application.SetCompatibleTextRenderingDefault(false);
		Application.Run(new Form1(centerStr, fontSize, fps));
	}
}

