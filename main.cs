using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Net;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

static class CommentServer {
	static List<WebSocket> webSocketClients = new List<WebSocket>();

	class Comment {
		int width, linePos, distFromRight;
		string text;

		public void move(int width) {
			distFromRight += width;
		}

		public int getDistFromRight() {
			return distFromRight;
		}

		public string getText() {
			return text;
		}

		public int getLinePos() {
			return linePos;
		}

		public int getWidth() {
			return width;
		}

		public void setWidth(int width) {
			this.width = width;
		}

		public Comment(string text, int linePos) {
			this.text = text;
			this.linePos = linePos;
			this.width = -1;
			distFromRight = 0;
		}
	}

	class Form1 : Form {
		public List<Comment> comments = new List<Comment>();
		const int FPS = 24;
		const int VISIBLE_SEC = 3;
		const int TOP_OFFSET = 20;
		int LINE_HEIGHT;
		Font FONT;

		PictureBox p;

		void keyDown(object sender, KeyEventArgs e) {
			if (e.KeyCode == Keys.Q) {
				this.Close();
				return;
			}

			PowerPoint.Application p;

			try {
				p = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
			} catch { 
				return;
			}

			if (p == null) {
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

		void drawComments() {
			var canvas = new Bitmap(p.Width, p.Height);
			var g = Graphics.FromImage(canvas);

			g.DrawString(
				"[CS]",
				FONT,
				Brushes.Black,
				0,
				p.Height - LINE_HEIGHT
			);

			lock(comments) {
				var finishedComments = new List<Comment>();
				foreach (var c in comments) {
					int cw = c.getWidth();
					if (cw == -1) {
						cw = TextRenderer.MeasureText(c.getText(), FONT).Width;
						c.setWidth(cw);
					}

					g.DrawString(
						c.getText(),
						FONT,
						Brushes.MediumSeaGreen,
						p.Width - c.getDistFromRight(),
						TOP_OFFSET + LINE_HEIGHT * c.getLinePos()
					);

					int moveWidth = (p.Width + cw) / (FPS * VISIBLE_SEC);
					c.move(moveWidth);

					if (p.Width + cw - c.getDistFromRight() < 0)
						finishedComments.Add(c);
				}

				foreach (var c in finishedComments)
					comments.Remove(c);
			}

			g.Dispose();
			p.Image = canvas;
		}

		void InitializeComponent() {
			var t = new Timer();
			t.Interval = 1000 / FPS;
			t.Start();
			t.Tick += (sender, e) => {
				drawComments();
			};

			// var t2 = new Timer();
			// t2.Interval = 1000;
			// t2.Start();
			// t2.Tick += (sender, e) => {
			// 	lock(comments) {
			// 		int pos = 0;
			// 		for (;;) {
			// 			bool isUsed = false;
			// 			foreach (var c in comments) {
			// 				if (pos == c.getLinePos()) {
			// 					isUsed = true;
			// 					break;
			// 				}
			// 			}
            //
			// 			if (!isUsed) {
			// 				break;
			// 			}
            //
			// 			pos++;
			// 		}
            //
			// 		comments.Add(new Comment("XX", pos));
			// 	}
			// };

			foreach (var s in Screen.AllScreens) {
				if (s.Primary) continue;

				Location = s.Bounds.Location;
				Size = s.Bounds.Size;
			}

			TopMost = true;
			StartPosition = FormStartPosition.Manual;
			KeyDown += keyDown;
			WindowState = FormWindowState.Maximized;
			TransparencyKey = BackColor;
			FormBorderStyle = FormBorderStyle.None;

			p = new PictureBox();
			p.Size = Size;
			Controls.Add(p);
		}

		public Form1() {
			FONT = new Font("Myrica M", 56, FontStyle.Bold);
			LINE_HEIGHT = TextRenderer.MeasureText("あ", FONT).Height;

			InitializeComponent();
			StartServer(this);
		}

	}

	static async void StartServer(Form1 f) {
		var listener = new HttpListener();
		listener.Prefixes.Add("http://+:6928/");
		listener.Start();
		for (;;) {
			var listenerContext = await listener.GetContextAsync();
			if (listenerContext.Request.IsWebSocketRequest) {
				ProcessRequest(f, listenerContext);
			} else {
				HttpListenerResponse res = listenerContext.Response;
				res.StatusCode = 200;
				byte[] html = File.ReadAllBytes("index.html");
				res.OutputStream.Write(html, 0, html.Length);
				res.Close();
			}
		}
	}

	static async void ProcessRequest(Form1 f, HttpListenerContext listenerContext) {
		var ws = (await listenerContext.AcceptWebSocketAsync(subProtocol:null)).WebSocket;

		webSocketClients.Add(ws);

		while (ws.State == WebSocketState.Open) {
			try {
				var buff = new ArraySegment<byte>(new byte[1024]);
				var ret = await ws.ReceiveAsync(buff, System.Threading.CancellationToken.None);

				if (ret.MessageType == WebSocketMessageType.Text) {
					Console.WriteLine(
						"{0}:String Received:{1}",
						DateTime.Now.ToString(),
						listenerContext.Request.RemoteEndPoint.Address.ToString()
					);
					Console.WriteLine(
						"Message={0}",
						Encoding.UTF8.GetString(buff.Take(ret.Count).ToArray())
					);

					lock(f.comments) {
						int pos = 0;
						for (;;) {
							bool isUsed = false;
							foreach (var c in f.comments) {
								if (pos == c.getLinePos()) {
									isUsed = true;
									break;
								}
							}
							if (!isUsed) break;
							pos++;
						}

						f.comments.Add(new Comment(
								Encoding.UTF8.GetString(buff.Take(ret.Count).ToArray()),
								pos
						));
					}
				} else if(ret.MessageType == WebSocketMessageType.Close) {
					Console.WriteLine(
						"{0}:Session Close:{1}",
						DateTime.Now.ToString(),
						listenerContext.Request.RemoteEndPoint.Address.ToString()
					);
					break;
				}
			} catch {
				Console.WriteLine(
					"{0}:Session Abort:{1}",
					DateTime.Now.ToString(),
					listenerContext.Request.RemoteEndPoint.Address.ToString()
				);

				break;
			}
		}
		webSocketClients.Remove(ws);
		ws.Dispose();
	}

	static void Main(string[] Args) {
		Application.EnableVisualStyles();
		Application.SetCompatibleTextRenderingDefault(false);
		Application.Run(new Form1());
	}
}


