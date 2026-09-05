import type { CSSProperties, SVGProps } from "react";
import {
  Ambulance,
  BarChart3,
  BriefcaseBusiness,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  Copy,
  Download,
  FileSpreadsheet,
  FileText,
  Grid2X2,
  HelpCircle,
  ImagePlus,
  Info,
  ListChecks,
  LogOut,
  LogIn,
  MapPin,
  Maximize2,
  Menu,
  MenuSquare,
  Minimize2,
  MoreHorizontal,
  MoreVertical,
  Pin,
  Pencil,
  ArrowDown,
  ArrowUp,
  ArrowUpDown,
  ArrowRight,
  ArrowLeft,
  SkipForward,
  Plane,
  Search,
  Settings,
  Share2,
  Shield,
  Eye,
  EyeOff,
  Swords,
  Upload,
  User,
  UserSearch,
  Users,
  Wrench,
  WrapText,
  XCircle,
  AlertTriangle,
} from "lucide-react";

type IconProps = SVGProps<SVGSVGElement> & {
  fontSize?: "inherit" | "small" | "medium" | "large";
  sx?: CSSProperties & Record<string, unknown>;
};

const sizeMap = {
  inherit: undefined,
  small: 18,
  medium: 22,
  large: 28,
} as const;

function withCompatIcon(Icon: typeof Shield) {
  return function CompatIcon({ fontSize = "medium", sx, style, ...props }: IconProps) {
    return (
      <Icon
        size={sizeMap[fontSize]}
        style={{ ...sx, ...style }}
        strokeWidth={1.8}
        {...props}
      />
    );
  };
}

export const AddPhotoAlternateOutlinedIcon = withCompatIcon(ImagePlus);
export const AnalyticsOutlinedIcon = withCompatIcon(BarChart3);
export const ArticleOutlinedIcon = withCompatIcon(FileText);
export const BeachAccessOutlinedIcon = withCompatIcon(Plane);
export const BuildOutlinedIcon = withCompatIcon(Wrench);
export const BusinessCenterOutlinedIcon = withCompatIcon(BriefcaseBusiness);
export const CalendarMonthOutlinedIcon = withCompatIcon(CalendarDays);
export const CheckCircleOutlineIcon = withCompatIcon(CheckCircle2);
export const CloudUploadOutlinedIcon = withCompatIcon(Upload);
export const ContentCopyOutlinedIcon = withCompatIcon(Copy);
export const DashboardOutlinedIcon = withCompatIcon(Grid2X2);
export const DeleteOutlineOutlinedIcon = withCompatIcon(XCircle);
export const EditOutlinedIcon = withCompatIcon(Pencil);
export const DescriptionOutlinedIcon = withCompatIcon(FileText);
export const ErrorOutlineOutlinedIcon = withCompatIcon(XCircle);
export const VisibilityOutlinedIcon = withCompatIcon(Eye);
export const VisibilityOffOutlinedIcon = withCompatIcon(EyeOff);
export const FileDownloadOutlinedIcon = withCompatIcon(Download);
export const FileUploadOutlinedIcon = withCompatIcon(Upload);
export const FormatListBulletedOutlinedIcon = withCompatIcon(ListChecks);
export const FullscreenExitOutlinedIcon = withCompatIcon(Minimize2);
export const FullscreenOutlinedIcon = withCompatIcon(Maximize2);
export const GridViewOutlinedIcon = withCompatIcon(Grid2X2);
export const GroupsOutlinedIcon = withCompatIcon(Users);
export const HelpOutlineOutlinedIcon = withCompatIcon(HelpCircle);
export const InfoOutlinedIcon = withCompatIcon(Info);
export const LocalHospitalOutlinedIcon = withCompatIcon(Ambulance);
export const LogoutOutlinedIcon = withCompatIcon(LogOut);
export const LoginOutlinedIcon = withCompatIcon(LogIn);
export const LocationOnOutlinedIcon = withCompatIcon(MapPin);
export const MenuOpenOutlinedIcon = withCompatIcon(MenuSquare);
export const MenuOutlinedIcon = withCompatIcon(Menu);
export const MilitaryTechOutlinedIcon = withCompatIcon(Swords);
export const MoreHorizOutlinedIcon = withCompatIcon(MoreHorizontal);
export const MoreVertOutlinedIcon = withCompatIcon(MoreVertical);
export const PersonOutlinedIcon = withCompatIcon(User);
export const PersonSearchOutlinedIcon = withCompatIcon(UserSearch);
export const PictureAsPdfOutlinedIcon = withCompatIcon(FileText);
export const PushPinOutlinedIcon = withCompatIcon(Pin);
export const SortArrowDownIcon = withCompatIcon(ArrowDown);
export const SortArrowUpIcon = withCompatIcon(ArrowUp);
export const SortArrowUpDownIcon = withCompatIcon(ArrowUpDown);
export const ArrowRightOutlinedIcon = withCompatIcon(ArrowRight);
export const ArrowLeftOutlinedIcon = withCompatIcon(ArrowLeft);
export const SkipNextOutlinedIcon = withCompatIcon(SkipForward);
export const SearchOutlinedIcon = withCompatIcon(Search);
export const SettingsOutlinedIcon = withCompatIcon(Settings);
export const ShareOutlinedIcon = withCompatIcon(Share2);
export const ShieldOutlinedIcon = withCompatIcon(Shield);
export const SyncAltOutlinedIcon = withCompatIcon(ClipboardList);
export const TableChartOutlinedIcon = withCompatIcon(FileSpreadsheet);
export const UploadFileOutlinedIcon = withCompatIcon(Upload);
export const WarningAmberOutlinedIcon = withCompatIcon(AlertTriangle);
export const WrapTextOutlinedIcon = withCompatIcon(WrapText);
