import './globals.css';

export const metadata = {
  title: 'DataGrid Excel Paste',
  description: 'Primary-key controlled data grid with paste, filter, group and sort',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
