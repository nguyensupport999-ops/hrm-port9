--
-- PostgreSQL database dump
--


-- Dumped from database version 17.6
-- Dumped by pg_dump version 18.4

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET transaction_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

--
-- Name: public; Type: SCHEMA; Schema: -; Owner: -
--

CREATE SCHEMA public;


--
-- Name: SCHEMA public; Type: COMMENT; Schema: -; Owner: -
--

COMMENT ON SCHEMA public IS 'standard public schema';


--
-- Name: date_add(date, integer); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.date_add(date date, days integer) RETURNS date
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN date + (days || ' days')::INTERVAL;

END;

$$;


--
-- Name: date_format(date, text); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.date_format(date_val date, format_str text) RETURNS text
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN TO_CHAR(date_val, 'YYYY-MM-DD');

END;

$$;


--
-- Name: date_sub(date, integer); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.date_sub(date date, days integer) RETURNS date
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN date - (days || ' days')::INTERVAL;

END;

$$;


--
-- Name: day(date); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.day(date_val date) RETURNS integer
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN EXTRACT(DAY FROM date_val);

END;

$$;


--
-- Name: get_next_birthday(date, date); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.get_next_birthday(birth_date date, reference_date date) RETURNS date
    LANGUAGE plpgsql IMMUTABLE
    AS $$

DECLARE

    this_year_birthday DATE;

BEGIN

    this_year_birthday := MAKE_DATE(EXTRACT(YEAR FROM reference_date)::INT, EXTRACT(MONTH FROM birth_date)::INT, EXTRACT(DAY FROM birth_date)::INT);

    IF this_year_birthday >= reference_date THEN

        RETURN this_year_birthday;

    ELSE

        RETURN MAKE_DATE(EXTRACT(YEAR FROM reference_date)::INT + 1, EXTRACT(MONTH FROM birth_date)::INT, EXTRACT(DAY FROM birth_date)::INT);

    END IF;

END;

$$;


--
-- Name: month(date); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.month(date_val date) RETURNS integer
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN EXTRACT(MONTH FROM date_val);

END;

$$;


--
-- Name: year(date); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.year(date_val date) RETURNS integer
    LANGUAGE plpgsql IMMUTABLE
    AS $$

BEGIN

    RETURN EXTRACT(YEAR FROM date_val);

END;

$$;


SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: cau_hinh_cong_van; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.cau_hinh_cong_van (
    id integer NOT NULL,
    loai character varying(20) NOT NULL,
    so_max integer DEFAULT 0 NOT NULL,
    prefix character varying(10),
    nam_hien_tai integer DEFAULT EXTRACT(year FROM CURRENT_DATE) NOT NULL,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: cau_hinh_cong_van_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.cau_hinh_cong_van_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: cau_hinh_cong_van_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.cau_hinh_cong_van_id_seq OWNED BY public.cau_hinh_cong_van.id;


--
-- Name: cau_hinh_he_thong; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.cau_hinh_he_thong (
    id integer NOT NULL,
    ten_cau_hinh character varying(50) NOT NULL,
    gia_tri character varying(100),
    mo_ta text,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: cau_hinh_he_thong_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.cau_hinh_he_thong_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: cau_hinh_he_thong_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.cau_hinh_he_thong_id_seq OWNED BY public.cau_hinh_he_thong.id;


--
-- Name: cham_cong; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.cham_cong (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    ngay date NOT NULL,
    ma_cong character varying(10),
    gio_tang_ca numeric(5,2) DEFAULT 0,
    gio_tang_ca_le numeric(5,2) DEFAULT 0,
    ghi_chu text,
    nguon character varying(20) DEFAULT 'THU_CONG'::character varying,
    created_by character varying(100),
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now(),
    ca_ngay character varying(10),
    ca_dem character varying(10)
);


--
-- Name: cham_cong_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.cham_cong_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: cham_cong_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.cham_cong_id_seq OWNED BY public.cham_cong.id;


--
-- Name: chat_messages; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.chat_messages (
    id integer NOT NULL,
    room_id integer,
    nhan_vien_id integer,
    noi_dung text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: chat_messages_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.chat_messages_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: chat_messages_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.chat_messages_id_seq OWNED BY public.chat_messages.id;


--
-- Name: chat_participants; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.chat_participants (
    id integer NOT NULL,
    room_id integer,
    user_id integer,
    joined_at timestamp without time zone DEFAULT now()
);


--
-- Name: chat_participants_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.chat_participants_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: chat_participants_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.chat_participants_id_seq OWNED BY public.chat_participants.id;


--
-- Name: chat_rooms; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.chat_rooms (
    id integer NOT NULL,
    ten_phong text NOT NULL,
    loai_phong character varying(20) DEFAULT 'NHOM'::character varying,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: chat_rooms_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.chat_rooms_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: chat_rooms_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.chat_rooms_id_seq OWNED BY public.chat_rooms.id;


--
-- Name: chuc_danh_ung_vien; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.chuc_danh_ung_vien (
    id integer NOT NULL,
    ten_chuc_danh character varying(150) NOT NULL,
    ghi_chu text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: chuc_danh_ung_vien_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.chuc_danh_ung_vien_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: chuc_danh_ung_vien_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.chuc_danh_ung_vien_id_seq OWNED BY public.chuc_danh_ung_vien.id;


--
-- Name: chuc_vu_danh_muc; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.chuc_vu_danh_muc (
    id integer NOT NULL,
    ten_chuc_vu character varying(150) NOT NULL,
    thu_tu integer DEFAULT 0,
    trang_thai character varying(20) DEFAULT 'Hoạt động'::character varying
);


--
-- Name: chuc_vu_danh_muc_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.chuc_vu_danh_muc_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: chuc_vu_danh_muc_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.chuc_vu_danh_muc_id_seq OWNED BY public.chuc_vu_danh_muc.id;


--
-- Name: cong_van_den; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.cong_van_den (
    id integer NOT NULL,
    so_cong_van character varying(50) NOT NULL,
    co_quan_phat_hanh character varying(200) NOT NULL,
    ngay_den date DEFAULT CURRENT_DATE NOT NULL,
    tieu_de text NOT NULL,
    trich_yeu text,
    file_url text,
    ghi_chu text,
    nguoi_tao character varying(100),
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now(),
    ma_vach_buu_dien character varying(100)
);


--
-- Name: cong_van_den_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.cong_van_den_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: cong_van_den_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.cong_van_den_id_seq OWNED BY public.cong_van_den.id;


--
-- Name: cong_van_di; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.cong_van_di (
    id integer NOT NULL,
    so_cong_van character varying(50) NOT NULL,
    phong_phat_hanh character varying(100) NOT NULL,
    ngay_phat_hanh date DEFAULT CURRENT_DATE NOT NULL,
    tieu_de text NOT NULL,
    trich_yeu text,
    file_url text,
    loai_cong_van character varying(20) NOT NULL,
    ghi_chu text,
    nguoi_tao character varying(100),
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now(),
    ma_vach_buu_dien character varying(100)
);


--
-- Name: cong_van_di_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.cong_van_di_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: cong_van_di_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.cong_van_di_id_seq OWNED BY public.cong_van_di.id;


--
-- Name: danh_muc_loai_cong_van; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_loai_cong_van (
    id integer NOT NULL,
    ma_loai character varying(10) NOT NULL,
    ten_loai character varying(50) NOT NULL,
    thu_tu integer DEFAULT 0,
    trang_thai boolean DEFAULT true
);


--
-- Name: danh_muc_loai_cong_van_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.danh_muc_loai_cong_van_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: danh_muc_loai_cong_van_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.danh_muc_loai_cong_van_id_seq OWNED BY public.danh_muc_loai_cong_van.id;


--
-- Name: danh_muc_loai_hop_dong; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_loai_hop_dong (
    id integer NOT NULL,
    ten_loai_hd character varying(255) NOT NULL,
    thu_tu integer DEFAULT 0,
    trang_thai character varying(20) DEFAULT 'active'::character varying,
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: danh_muc_loai_hop_dong_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.danh_muc_loai_hop_dong_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: danh_muc_loai_hop_dong_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.danh_muc_loai_hop_dong_id_seq OWNED BY public.danh_muc_loai_hop_dong.id;


--
-- Name: danh_muc_phong_ban; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_phong_ban (
    id integer NOT NULL,
    ten_phong_ban character varying(255) NOT NULL,
    thu_tu integer DEFAULT 0,
    trang_thai character varying(20) DEFAULT 'active'::character varying,
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: danh_muc_phong_ban_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.danh_muc_phong_ban_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: danh_muc_phong_ban_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.danh_muc_phong_ban_id_seq OWNED BY public.danh_muc_phong_ban.id;


--
-- Name: danh_muc_phuong_xa; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_phuong_xa (
    ma_xa character varying(10) NOT NULL,
    ten_xa character varying(100) NOT NULL,
    ma_tinh character varying(10)
);


--
-- Name: danh_muc_tinh; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_tinh (
    ma_tinh character varying(10) NOT NULL,
    ten_tinh character varying(100) NOT NULL
);


--
-- Name: danh_muc_trinh_do_hoc_van; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.danh_muc_trinh_do_hoc_van (
    id integer NOT NULL,
    ten_trinh_do character varying(255) NOT NULL,
    thu_tu integer DEFAULT 0,
    trang_thai character varying(20) DEFAULT 'active'::character varying,
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: danh_muc_trinh_do_hoc_van_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.danh_muc_trinh_do_hoc_van_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: danh_muc_trinh_do_hoc_van_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.danh_muc_trinh_do_hoc_van_id_seq OWNED BY public.danh_muc_trinh_do_hoc_van.id;


--
-- Name: ho_so_nhan_vien; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.ho_so_nhan_vien (
    id bigint NOT NULL,
    nhan_vien_id bigint,
    loai_ho_so character varying(100),
    ten_file character varying(255),
    duong_dan_file character varying(500),
    ngay_upload date,
    ghi_chu text
);


--
-- Name: ho_so_nhan_vien_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.ho_so_nhan_vien_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: ho_so_nhan_vien_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.ho_so_nhan_vien_id_seq OWNED BY public.ho_so_nhan_vien.id;


--
-- Name: hop_dong_kinh_te; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.hop_dong_kinh_te (
    id integer NOT NULL,
    so_hop_dong character varying(50) NOT NULL,
    ten_doi_tac character varying(200) NOT NULL,
    ngay_ky date DEFAULT CURRENT_DATE NOT NULL,
    trich_yeu text,
    file_url text,
    ghi_chu text,
    nguoi_tao character varying(100),
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: hop_dong_kinh_te_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.hop_dong_kinh_te_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: hop_dong_kinh_te_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.hop_dong_kinh_te_id_seq OWNED BY public.hop_dong_kinh_te.id;


--
-- Name: lich_su_cong_tac; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.lich_su_cong_tac (
    id bigint NOT NULL,
    nhan_vien_id bigint,
    tu_ngay date NOT NULL,
    den_ngay date,
    chuc_danh character varying(100),
    phong_ban character varying(100),
    noi_lam_viec character varying(200),
    loai_hop_dong character varying(50),
    he_so_luong numeric(10,2),
    ghi_chu text,
    so_hop_dong text,
    created_at timestamp without time zone DEFAULT CURRENT_TIMESTAMP
);


--
-- Name: lich_su_cong_tac_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.lich_su_cong_tac_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: lich_su_cong_tac_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.lich_su_cong_tac_id_seq OWNED BY public.lich_su_cong_tac.id;


--
-- Name: lich_su_gui_loi_chuc; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.lich_su_gui_loi_chuc (
    id integer NOT NULL,
    nhan_vien_id integer,
    loai_chuc character varying(50) DEFAULT 'SINH_NHAT'::character varying,
    noi_dung text,
    kenh_gui character varying(20),
    trang_thai character varying(20),
    ngay_gui timestamp without time zone DEFAULT CURRENT_TIMESTAMP,
    created_at timestamp without time zone DEFAULT CURRENT_TIMESTAMP
);


--
-- Name: lich_su_gui_loi_chuc_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.lich_su_gui_loi_chuc_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: lich_su_gui_loi_chuc_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.lich_su_gui_loi_chuc_id_seq OWNED BY public.lich_su_gui_loi_chuc.id;


--
-- Name: mau_dieu_hop_dong; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.mau_dieu_hop_dong (
    id integer NOT NULL,
    loai_hd character varying(10) NOT NULL,
    ma_dieu character varying(30) NOT NULL,
    tieu_de text,
    noi_dung text,
    thu_tu integer DEFAULT 0,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: mau_dieu_hop_dong_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.mau_dieu_hop_dong_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: mau_dieu_hop_dong_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.mau_dieu_hop_dong_id_seq OWNED BY public.mau_dieu_hop_dong.id;


--
-- Name: nhan_vien; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.nhan_vien (
    id bigint NOT NULL,
    stt integer,
    ma_nv character varying(50) NOT NULL,
    so_hdld character varying(50),
    ho_ten character varying(255) NOT NULL,
    chuc_danh_nghe character varying(255),
    ngay_sinh date,
    gioi_tinh character varying(10),
    tinh_trang_hon_nhan character varying(20),
    so_cccd character varying(20),
    ngay_cap_cccd date,
    noi_cap_cccd character varying(255),
    nguyen_quan character varying(500),
    thuong_tru character varying(500),
    dien_thoai character varying(20),
    email character varying(255),
    email_lien_he character varying(255),
    ho_so character varying(10),
    luong_bao_hiem character varying(100),
    ma_so_bhxh character varying(50),
    ngay_vao_lam date,
    noi_lam_viec character varying(255),
    so_tai_khoan_nh character varying(50),
    chi_nhanh_nh character varying(255),
    ngay_ky_hd date,
    thoi_han_hd character varying(255),
    loai_hop_dong character varying(50),
    nhom_bhxh character varying(50),
    thang_bat_dau_bh date,
    thang_ket_thuc_bh date,
    ghi_chu text,
    trang_thai character varying(30) DEFAULT 'DANG_LAM'::character varying,
    trang_thai_bhxh character varying(30) DEFAULT 'DANG_DONG'::character varying,
    phong_ban_lam_viec character varying(255),
    ngay_ket_thuc date,
    ly_do_nghi character varying(500),
    quoc_tich character varying(50) DEFAULT 'Việt Nam'::character varying,
    dan_toc character varying(50) DEFAULT 'Kinh'::character varying,
    he_so_luong numeric(10,2),
    phu_cap_chuc_vu numeric(10,2),
    phu_cap_tnvk numeric(5,2),
    phu_cap_tnn numeric(5,2),
    muc_huong_bhyt character varying(10) DEFAULT '80%'::character varying,
    ty_le_dong numeric(5,2),
    muc_tien_dong numeric(15,2),
    phuong_thuc_dong character varying(50) DEFAULT 'Hàng tháng'::character varying,
    tinh_nhan_hs character varying(255),
    phuong_nhan_hs character varying(255),
    dia_chi_nhan_hs text,
    tinh_kcb character varying(255),
    noi_dang_ky_kcb character varying(255),
    dang_ky_nhan_so character varying(10) DEFAULT 'Có'::character varying,
    ho_ten_chu_ho character varying(100),
    so_cccd_chu_ho character varying(20),
    tinh_thanh_pho_chu_ho character varying(100),
    phuong_xa_chu_ho character varying(100),
    tinh_thanh_pho_thuong_tru character varying(100),
    ma_tinh_thuong_tru character varying(10),
    phuong_xa_thuong_tru character varying(100),
    ma_phuong_xa_thuong_tru character varying(10),
    ngay_chinh_thuc date,
    da_truy_nguon_bhxh integer DEFAULT 0,
    vi_tri_id integer,
    ten_don_vi_thu_huong character varying(255),
    mat_khau_hash text,
    phai_doi_mat_khau boolean DEFAULT false,
    vai_tro character varying(20) DEFAULT 'nhan_vien'::character varying,
    ten_dang_nhap text,
    trinh_do text,
    anh_ho_so text,
    chuc_vu character varying(100),
    ngay_qd_ns date,
    so_luong_npt integer DEFAULT 0
);


--
-- Name: nhan_vien_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.nhan_vien_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: nhan_vien_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.nhan_vien_id_seq OWNED BY public.nhan_vien.id;


--
-- Name: phu_luc_gia_dinh; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.phu_luc_gia_dinh (
    id bigint NOT NULL,
    nhan_vien_id bigint,
    ho_ten character varying(100) NOT NULL,
    ngay_sinh date,
    gioi_tinh character varying(10),
    quoc_tich character varying(50) DEFAULT 'Việt Nam'::character varying,
    dan_toc character varying(50) DEFAULT 'Kinh'::character varying,
    quan_he_voi_chu_ho character varying(50),
    tinh_thanh_pho character varying(100),
    phuong_xa character varying(100)
);


--
-- Name: phu_luc_gia_dinh_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.phu_luc_gia_dinh_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: phu_luc_gia_dinh_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.phu_luc_gia_dinh_id_seq OWNED BY public.phu_luc_gia_dinh.id;


--
-- Name: quyet_dinh_nhan_su; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.quyet_dinh_nhan_su (
    id bigint NOT NULL,
    nhan_vien_id bigint,
    loai_quyet_dinh character varying(50),
    so_quyet_dinh character varying(50),
    ngay_quyet_dinh date NOT NULL,
    ngay_hieu_luc date NOT NULL,
    noi_dung text,
    ghi_chu text,
    trang_thai character varying(20) DEFAULT 'CO_HIEU_LUC'::character varying,
    nguoi_ky character varying(100),
    chuc_danh_cu character varying(100),
    chuc_danh_moi character varying(100),
    phong_ban_cu character varying(100),
    phong_ban_moi character varying(100),
    loai_hop_dong_cu character varying(50),
    loai_hop_dong_moi character varying(50),
    he_so_luong_cu numeric(10,2),
    he_so_luong_moi numeric(10,2),
    created_at timestamp without time zone DEFAULT CURRENT_TIMESTAMP,
    so_hd_cu character varying(50),
    so_hd_moi character varying(50),
    gia_tri_truoc character varying(150),
    gia_tri_sau character varying(150),
    file_url text,
    nguoi_tao character varying(100),
    loai_qd character varying(30),
    so_qd character varying(50),
    ngay_qd date DEFAULT CURRENT_DATE
);


--
-- Name: quyet_dinh_nhan_su_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.quyet_dinh_nhan_su_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: quyet_dinh_nhan_su_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.quyet_dinh_nhan_su_id_seq OWNED BY public.quyet_dinh_nhan_su.id;


--
-- Name: tp_bsc_strategy_map; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_bsc_strategy_map (
    id integer NOT NULL,
    khia_canh text NOT NULL,
    muc_tieu text NOT NULL,
    chi_so_do_luong text,
    chi_tieu text,
    phong_ban_phu_trach text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_bsc_strategy_map_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_bsc_strategy_map_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_bsc_strategy_map_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_bsc_strategy_map_id_seq OWNED BY public.tp_bsc_strategy_map.id;


--
-- Name: tp_employee_dependents; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_employee_dependents (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    ho_ten text NOT NULL,
    quan_he text,
    ngay_sinh date,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_employee_dependents_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_employee_dependents_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_employee_dependents_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_employee_dependents_id_seq OWNED BY public.tp_employee_dependents.id;


--
-- Name: tp_employee_p1_assignment; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_employee_p1_assignment (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    ma_chuc_danh text,
    ma_ngach text NOT NULL,
    bac integer NOT NULL,
    hieu_luc_tu date DEFAULT CURRENT_DATE,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_employee_p1_assignment_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_employee_p1_assignment_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_employee_p1_assignment_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_employee_p1_assignment_id_seq OWNED BY public.tp_employee_p1_assignment.id;


--
-- Name: tp_employee_p2_score; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_employee_p2_score (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    thang integer NOT NULL,
    nam integer NOT NULL,
    diem_nang_luc numeric(4,2) NOT NULL,
    nguoi_danh_gia text,
    ghi_chu text,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_employee_p2_score_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_employee_p2_score_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_employee_p2_score_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_employee_p2_score_id_seq OWNED BY public.tp_employee_p2_score.id;


--
-- Name: tp_employee_p3_score; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_employee_p3_score (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    thang integer NOT NULL,
    nam integer NOT NULL,
    ty_le_hoan_thanh numeric(6,2) NOT NULL,
    ghi_chu text,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_employee_p3_score_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_employee_p3_score_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_employee_p3_score_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_employee_p3_score_id_seq OWNED BY public.tp_employee_p3_score.id;


--
-- Name: tp_functional_matrix; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_functional_matrix (
    id integer NOT NULL,
    phong_ban text NOT NULL,
    nhom_chuc_nang text NOT NULL,
    muc_do_tham_gia text,
    bsc_id integer,
    ghi_chu text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_functional_matrix_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_functional_matrix_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_functional_matrix_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_functional_matrix_id_seq OWNED BY public.tp_functional_matrix.id;


--
-- Name: tp_job_description; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_job_description (
    id integer NOT NULL,
    ma_chuc_danh text NOT NULL,
    ten_chuc_danh text NOT NULL,
    phong_ban text,
    bao_cao_cho text,
    muc_tieu_cong_viec text,
    nhiem_vu_chinh text,
    yeu_cau_trinh_do text,
    yeu_cau_kinh_nghiem text,
    yeu_cau_ky_nang text,
    dieu_kien_lam_viec text,
    trang_thai text DEFAULT 'active'::text,
    created_at timestamp without time zone DEFAULT now(),
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_job_description_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_job_description_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_job_description_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_job_description_id_seq OWNED BY public.tp_job_description.id;


--
-- Name: tp_job_evaluation; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_job_evaluation (
    id integer NOT NULL,
    ma_chuc_danh text,
    diem_chi_tiet jsonb,
    tong_diem numeric(6,2),
    ma_ngach text,
    ten_ngach text,
    nguoi_danh_gia text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_job_evaluation_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_job_evaluation_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_job_evaluation_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_job_evaluation_id_seq OWNED BY public.tp_job_evaluation.id;


--
-- Name: tp_p1_salary_scale; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_p1_salary_scale (
    id integer NOT NULL,
    ma_ngach text NOT NULL,
    ten_ngach text,
    bac integer NOT NULL,
    he_so numeric(6,3),
    luong_p1 numeric(14,0) NOT NULL,
    ghi_chu text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_p1_salary_scale_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_p1_salary_scale_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_p1_salary_scale_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_p1_salary_scale_id_seq OWNED BY public.tp_p1_salary_scale.id;


--
-- Name: tp_p2_competency; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_p2_competency (
    id integer NOT NULL,
    nhom_nang_luc text NOT NULL,
    ten_nang_luc text NOT NULL,
    mo_ta text,
    ap_dung_chuc_danh text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_p2_competency_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_p2_competency_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_p2_competency_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_p2_competency_id_seq OWNED BY public.tp_p2_competency.id;


--
-- Name: tp_p3_kpi_system; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_p3_kpi_system (
    id integer NOT NULL,
    ma_kpi text NOT NULL,
    ten_kpi text NOT NULL,
    phong_ban text,
    khia_canh_bsc text,
    don_vi_tinh text,
    trong_so numeric(5,2),
    tan_suat text DEFAULT 'Tháng'::text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_p3_kpi_system_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_p3_kpi_system_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_p3_kpi_system_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_p3_kpi_system_id_seq OWNED BY public.tp_p3_kpi_system.id;


--
-- Name: tp_payroll_config; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_payroll_config (
    id integer NOT NULL,
    cau_hinh jsonb NOT NULL,
    updated_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_payroll_config_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_payroll_config_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_payroll_config_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_payroll_config_id_seq OWNED BY public.tp_payroll_config.id;


--
-- Name: tp_payroll_detail; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_payroll_detail (
    id integer NOT NULL,
    period_id integer,
    nhan_vien_id integer NOT NULL,
    p1_luong_vi_tri numeric(14,0) DEFAULT 0,
    p2_luong_nang_luc numeric(14,0) DEFAULT 0,
    p3_luong_hieu_qua numeric(14,0) DEFAULT 0,
    phu_cap_chuc_vu numeric(14,0) DEFAULT 0,
    phu_cap_tham_nien numeric(14,0) DEFAULT 0,
    phu_cap_trach_nhiem numeric(14,0) DEFAULT 0,
    phu_cap_khac numeric(14,0) DEFAULT 0,
    tong_thu_nhap numeric(14,0) DEFAULT 0,
    luong_dong_bh numeric(14,0) DEFAULT 0,
    bhxh_nld numeric(14,0) DEFAULT 0,
    bhyt_nld numeric(14,0) DEFAULT 0,
    bhtn_nld numeric(14,0) DEFAULT 0,
    doan_phi numeric(14,0) DEFAULT 0,
    so_nguoi_phu_thuoc integer DEFAULT 0,
    giam_tru_gia_canh numeric(14,0) DEFAULT 0,
    thu_nhap_tinh_thue numeric(14,0) DEFAULT 0,
    thue_tncn numeric(14,0) DEFAULT 0,
    tong_khau_tru numeric(14,0) DEFAULT 0,
    thuc_nhan numeric(14,0) DEFAULT 0,
    chi_tiet jsonb,
    ngay_tinh timestamp without time zone DEFAULT now(),
    nguoi_tinh text,
    da_gui_chat boolean DEFAULT false
);


--
-- Name: tp_payroll_detail_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_payroll_detail_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_payroll_detail_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_payroll_detail_id_seq OWNED BY public.tp_payroll_detail.id;


--
-- Name: tp_payroll_period; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_payroll_period (
    id integer NOT NULL,
    thang integer NOT NULL,
    nam integer NOT NULL,
    trang_thai text DEFAULT 'draft'::text,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: tp_payroll_period_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_payroll_period_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_payroll_period_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_payroll_period_id_seq OWNED BY public.tp_payroll_period.id;


--
-- Name: tp_policy_version; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tp_policy_version (
    id integer NOT NULL,
    phien_ban integer NOT NULL,
    ngay_xuat_ban timestamp without time zone DEFAULT now(),
    nguoi_xuat_ban text,
    trang_thai text DEFAULT 'published'::text,
    thong_ke jsonb,
    ghi_chu text
);


--
-- Name: tp_policy_version_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tp_policy_version_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tp_policy_version_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tp_policy_version_id_seq OWNED BY public.tp_policy_version.id;


--
-- Name: ung_vien; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.ung_vien (
    id bigint,
    ma_uv character varying(50),
    ho_ten character varying(255),
    vi_tri_du_tuyen character varying(255),
    ngay_sinh date,
    gioi_tinh character varying(10),
    dien_thoai character varying(20),
    ngay_vao_lam date,
    luong_bao_hiem text,
    ma_nv character varying(50),
    trang_thai character varying(30),
    created_at timestamp without time zone,
    chuc_danh_nghe character varying(255),
    tinh_trang_hon_nhan character varying(20),
    so_cccd character varying(20),
    ngay_cap_cccd date,
    noi_cap_cccd character varying(255),
    nguyen_quan character varying(500),
    thuong_tru character varying(500),
    email character varying(255),
    ho_so character varying(10),
    ma_so_bhxh character varying(50),
    noi_lam_viec character varying(255),
    so_tai_khoan_nh character varying(50),
    chi_nhanh_nh character varying(255),
    ngay_ky_hd date,
    thoi_han_hd character varying(255),
    ghi_chu text
);


--
-- Name: ung_vien_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.ung_vien_id_seq
    START WITH 215
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: ung_vien_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.ung_vien_id_seq OWNED BY public.ung_vien.id;


--
-- Name: vi_tri_cong_tac; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.vi_tri_cong_tac (
    id bigint NOT NULL,
    ten_vi_tri character varying(255) NOT NULL,
    so_luong_can_tuyen integer DEFAULT 0,
    phu_cap_tang_ca integer DEFAULT 0,
    phu_cap_ca_dem integer DEFAULT 0,
    ghi_chu text
);


--
-- Name: vi_tri_cong_tac_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.vi_tri_cong_tac_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: vi_tri_cong_tac_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.vi_tri_cong_tac_id_seq OWNED BY public.vi_tri_cong_tac.id;


--
-- Name: yeu_cau_reset_mk; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.yeu_cau_reset_mk (
    id integer NOT NULL,
    nhan_vien_id integer NOT NULL,
    otp_code character varying(10) NOT NULL,
    het_han timestamp without time zone NOT NULL,
    da_dung boolean DEFAULT false,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: yeu_cau_reset_mk_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.yeu_cau_reset_mk_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: yeu_cau_reset_mk_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.yeu_cau_reset_mk_id_seq OWNED BY public.yeu_cau_reset_mk.id;


--
-- Name: cau_hinh_cong_van id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_cong_van ALTER COLUMN id SET DEFAULT nextval('public.cau_hinh_cong_van_id_seq'::regclass);


--
-- Name: cau_hinh_he_thong id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_he_thong ALTER COLUMN id SET DEFAULT nextval('public.cau_hinh_he_thong_id_seq'::regclass);


--
-- Name: cham_cong id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cham_cong ALTER COLUMN id SET DEFAULT nextval('public.cham_cong_id_seq'::regclass);


--
-- Name: chat_messages id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_messages ALTER COLUMN id SET DEFAULT nextval('public.chat_messages_id_seq'::regclass);


--
-- Name: chat_participants id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_participants ALTER COLUMN id SET DEFAULT nextval('public.chat_participants_id_seq'::regclass);


--
-- Name: chat_rooms id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_rooms ALTER COLUMN id SET DEFAULT nextval('public.chat_rooms_id_seq'::regclass);


--
-- Name: chuc_danh_ung_vien id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_danh_ung_vien ALTER COLUMN id SET DEFAULT nextval('public.chuc_danh_ung_vien_id_seq'::regclass);


--
-- Name: chuc_vu_danh_muc id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_vu_danh_muc ALTER COLUMN id SET DEFAULT nextval('public.chuc_vu_danh_muc_id_seq'::regclass);


--
-- Name: cong_van_den id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cong_van_den ALTER COLUMN id SET DEFAULT nextval('public.cong_van_den_id_seq'::regclass);


--
-- Name: cong_van_di id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cong_van_di ALTER COLUMN id SET DEFAULT nextval('public.cong_van_di_id_seq'::regclass);


--
-- Name: danh_muc_loai_cong_van id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_cong_van ALTER COLUMN id SET DEFAULT nextval('public.danh_muc_loai_cong_van_id_seq'::regclass);


--
-- Name: danh_muc_loai_hop_dong id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_hop_dong ALTER COLUMN id SET DEFAULT nextval('public.danh_muc_loai_hop_dong_id_seq'::regclass);


--
-- Name: danh_muc_phong_ban id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_phong_ban ALTER COLUMN id SET DEFAULT nextval('public.danh_muc_phong_ban_id_seq'::regclass);


--
-- Name: danh_muc_trinh_do_hoc_van id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_trinh_do_hoc_van ALTER COLUMN id SET DEFAULT nextval('public.danh_muc_trinh_do_hoc_van_id_seq'::regclass);


--
-- Name: ho_so_nhan_vien id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.ho_so_nhan_vien ALTER COLUMN id SET DEFAULT nextval('public.ho_so_nhan_vien_id_seq'::regclass);


--
-- Name: hop_dong_kinh_te id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.hop_dong_kinh_te ALTER COLUMN id SET DEFAULT nextval('public.hop_dong_kinh_te_id_seq'::regclass);


--
-- Name: lich_su_cong_tac id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_cong_tac ALTER COLUMN id SET DEFAULT nextval('public.lich_su_cong_tac_id_seq'::regclass);


--
-- Name: lich_su_gui_loi_chuc id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_gui_loi_chuc ALTER COLUMN id SET DEFAULT nextval('public.lich_su_gui_loi_chuc_id_seq'::regclass);


--
-- Name: mau_dieu_hop_dong id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.mau_dieu_hop_dong ALTER COLUMN id SET DEFAULT nextval('public.mau_dieu_hop_dong_id_seq'::regclass);


--
-- Name: nhan_vien id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.nhan_vien ALTER COLUMN id SET DEFAULT nextval('public.nhan_vien_id_seq'::regclass);


--
-- Name: phu_luc_gia_dinh id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.phu_luc_gia_dinh ALTER COLUMN id SET DEFAULT nextval('public.phu_luc_gia_dinh_id_seq'::regclass);


--
-- Name: quyet_dinh_nhan_su id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.quyet_dinh_nhan_su ALTER COLUMN id SET DEFAULT nextval('public.quyet_dinh_nhan_su_id_seq'::regclass);


--
-- Name: tp_bsc_strategy_map id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_bsc_strategy_map ALTER COLUMN id SET DEFAULT nextval('public.tp_bsc_strategy_map_id_seq'::regclass);


--
-- Name: tp_employee_dependents id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_dependents ALTER COLUMN id SET DEFAULT nextval('public.tp_employee_dependents_id_seq'::regclass);


--
-- Name: tp_employee_p1_assignment id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p1_assignment ALTER COLUMN id SET DEFAULT nextval('public.tp_employee_p1_assignment_id_seq'::regclass);


--
-- Name: tp_employee_p2_score id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p2_score ALTER COLUMN id SET DEFAULT nextval('public.tp_employee_p2_score_id_seq'::regclass);


--
-- Name: tp_employee_p3_score id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p3_score ALTER COLUMN id SET DEFAULT nextval('public.tp_employee_p3_score_id_seq'::regclass);


--
-- Name: tp_functional_matrix id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_functional_matrix ALTER COLUMN id SET DEFAULT nextval('public.tp_functional_matrix_id_seq'::regclass);


--
-- Name: tp_job_description id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_description ALTER COLUMN id SET DEFAULT nextval('public.tp_job_description_id_seq'::regclass);


--
-- Name: tp_job_evaluation id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_evaluation ALTER COLUMN id SET DEFAULT nextval('public.tp_job_evaluation_id_seq'::regclass);


--
-- Name: tp_p1_salary_scale id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p1_salary_scale ALTER COLUMN id SET DEFAULT nextval('public.tp_p1_salary_scale_id_seq'::regclass);


--
-- Name: tp_p2_competency id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p2_competency ALTER COLUMN id SET DEFAULT nextval('public.tp_p2_competency_id_seq'::regclass);


--
-- Name: tp_p3_kpi_system id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p3_kpi_system ALTER COLUMN id SET DEFAULT nextval('public.tp_p3_kpi_system_id_seq'::regclass);


--
-- Name: tp_payroll_config id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_config ALTER COLUMN id SET DEFAULT nextval('public.tp_payroll_config_id_seq'::regclass);


--
-- Name: tp_payroll_detail id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_detail ALTER COLUMN id SET DEFAULT nextval('public.tp_payroll_detail_id_seq'::regclass);


--
-- Name: tp_payroll_period id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_period ALTER COLUMN id SET DEFAULT nextval('public.tp_payroll_period_id_seq'::regclass);


--
-- Name: tp_policy_version id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_policy_version ALTER COLUMN id SET DEFAULT nextval('public.tp_policy_version_id_seq'::regclass);


--
-- Name: ung_vien id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.ung_vien ALTER COLUMN id SET DEFAULT nextval('public.ung_vien_id_seq'::regclass);


--
-- Name: vi_tri_cong_tac id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.vi_tri_cong_tac ALTER COLUMN id SET DEFAULT nextval('public.vi_tri_cong_tac_id_seq'::regclass);


--
-- Name: yeu_cau_reset_mk id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.yeu_cau_reset_mk ALTER COLUMN id SET DEFAULT nextval('public.yeu_cau_reset_mk_id_seq'::regclass);


--
-- Name: cau_hinh_cong_van cau_hinh_cong_van_loai_nam_hien_tai_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_cong_van
    ADD CONSTRAINT cau_hinh_cong_van_loai_nam_hien_tai_key UNIQUE (loai, nam_hien_tai);


--
-- Name: cau_hinh_cong_van cau_hinh_cong_van_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_cong_van
    ADD CONSTRAINT cau_hinh_cong_van_pkey PRIMARY KEY (id);


--
-- Name: cau_hinh_he_thong cau_hinh_he_thong_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_he_thong
    ADD CONSTRAINT cau_hinh_he_thong_pkey PRIMARY KEY (id);


--
-- Name: cau_hinh_he_thong cau_hinh_he_thong_ten_cau_hinh_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cau_hinh_he_thong
    ADD CONSTRAINT cau_hinh_he_thong_ten_cau_hinh_key UNIQUE (ten_cau_hinh);


--
-- Name: cham_cong cham_cong_nhan_vien_id_ngay_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cham_cong
    ADD CONSTRAINT cham_cong_nhan_vien_id_ngay_key UNIQUE (nhan_vien_id, ngay);


--
-- Name: cham_cong cham_cong_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cham_cong
    ADD CONSTRAINT cham_cong_pkey PRIMARY KEY (id);


--
-- Name: chat_messages chat_messages_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_messages
    ADD CONSTRAINT chat_messages_pkey PRIMARY KEY (id);


--
-- Name: chat_participants chat_participants_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_participants
    ADD CONSTRAINT chat_participants_pkey PRIMARY KEY (id);


--
-- Name: chat_participants chat_participants_room_id_user_id_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_participants
    ADD CONSTRAINT chat_participants_room_id_user_id_key UNIQUE (room_id, user_id);


--
-- Name: chat_rooms chat_rooms_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_rooms
    ADD CONSTRAINT chat_rooms_pkey PRIMARY KEY (id);


--
-- Name: chuc_danh_ung_vien chuc_danh_ung_vien_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_danh_ung_vien
    ADD CONSTRAINT chuc_danh_ung_vien_pkey PRIMARY KEY (id);


--
-- Name: chuc_danh_ung_vien chuc_danh_ung_vien_ten_chuc_danh_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_danh_ung_vien
    ADD CONSTRAINT chuc_danh_ung_vien_ten_chuc_danh_key UNIQUE (ten_chuc_danh);


--
-- Name: chuc_vu_danh_muc chuc_vu_danh_muc_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_vu_danh_muc
    ADD CONSTRAINT chuc_vu_danh_muc_pkey PRIMARY KEY (id);


--
-- Name: chuc_vu_danh_muc chuc_vu_danh_muc_ten_chuc_vu_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chuc_vu_danh_muc
    ADD CONSTRAINT chuc_vu_danh_muc_ten_chuc_vu_key UNIQUE (ten_chuc_vu);


--
-- Name: cong_van_den cong_van_den_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cong_van_den
    ADD CONSTRAINT cong_van_den_pkey PRIMARY KEY (id);


--
-- Name: cong_van_di cong_van_di_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cong_van_di
    ADD CONSTRAINT cong_van_di_pkey PRIMARY KEY (id);


--
-- Name: danh_muc_loai_cong_van danh_muc_loai_cong_van_ma_loai_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_cong_van
    ADD CONSTRAINT danh_muc_loai_cong_van_ma_loai_key UNIQUE (ma_loai);


--
-- Name: danh_muc_loai_cong_van danh_muc_loai_cong_van_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_cong_van
    ADD CONSTRAINT danh_muc_loai_cong_van_pkey PRIMARY KEY (id);


--
-- Name: danh_muc_loai_hop_dong danh_muc_loai_hop_dong_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_hop_dong
    ADD CONSTRAINT danh_muc_loai_hop_dong_pkey PRIMARY KEY (id);


--
-- Name: danh_muc_loai_hop_dong danh_muc_loai_hop_dong_ten_loai_hd_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_loai_hop_dong
    ADD CONSTRAINT danh_muc_loai_hop_dong_ten_loai_hd_key UNIQUE (ten_loai_hd);


--
-- Name: danh_muc_phong_ban danh_muc_phong_ban_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_phong_ban
    ADD CONSTRAINT danh_muc_phong_ban_pkey PRIMARY KEY (id);


--
-- Name: danh_muc_phong_ban danh_muc_phong_ban_ten_phong_ban_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_phong_ban
    ADD CONSTRAINT danh_muc_phong_ban_ten_phong_ban_key UNIQUE (ten_phong_ban);


--
-- Name: danh_muc_phuong_xa danh_muc_phuong_xa_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_phuong_xa
    ADD CONSTRAINT danh_muc_phuong_xa_pkey PRIMARY KEY (ma_xa);


--
-- Name: danh_muc_tinh danh_muc_tinh_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_tinh
    ADD CONSTRAINT danh_muc_tinh_pkey PRIMARY KEY (ma_tinh);


--
-- Name: danh_muc_trinh_do_hoc_van danh_muc_trinh_do_hoc_van_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_trinh_do_hoc_van
    ADD CONSTRAINT danh_muc_trinh_do_hoc_van_pkey PRIMARY KEY (id);


--
-- Name: danh_muc_trinh_do_hoc_van danh_muc_trinh_do_hoc_van_ten_trinh_do_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_trinh_do_hoc_van
    ADD CONSTRAINT danh_muc_trinh_do_hoc_van_ten_trinh_do_key UNIQUE (ten_trinh_do);


--
-- Name: ho_so_nhan_vien ho_so_nhan_vien_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.ho_so_nhan_vien
    ADD CONSTRAINT ho_so_nhan_vien_pkey PRIMARY KEY (id);


--
-- Name: hop_dong_kinh_te hop_dong_kinh_te_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.hop_dong_kinh_te
    ADD CONSTRAINT hop_dong_kinh_te_pkey PRIMARY KEY (id);


--
-- Name: lich_su_cong_tac lich_su_cong_tac_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_cong_tac
    ADD CONSTRAINT lich_su_cong_tac_pkey PRIMARY KEY (id);


--
-- Name: lich_su_gui_loi_chuc lich_su_gui_loi_chuc_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_gui_loi_chuc
    ADD CONSTRAINT lich_su_gui_loi_chuc_pkey PRIMARY KEY (id);


--
-- Name: mau_dieu_hop_dong mau_dieu_hop_dong_loai_hd_ma_dieu_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.mau_dieu_hop_dong
    ADD CONSTRAINT mau_dieu_hop_dong_loai_hd_ma_dieu_key UNIQUE (loai_hd, ma_dieu);


--
-- Name: mau_dieu_hop_dong mau_dieu_hop_dong_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.mau_dieu_hop_dong
    ADD CONSTRAINT mau_dieu_hop_dong_pkey PRIMARY KEY (id);


--
-- Name: nhan_vien nhan_vien_dien_thoai_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.nhan_vien
    ADD CONSTRAINT nhan_vien_dien_thoai_key UNIQUE (dien_thoai);


--
-- Name: nhan_vien nhan_vien_ma_nv_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.nhan_vien
    ADD CONSTRAINT nhan_vien_ma_nv_key UNIQUE (ma_nv);


--
-- Name: nhan_vien nhan_vien_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.nhan_vien
    ADD CONSTRAINT nhan_vien_pkey PRIMARY KEY (id);


--
-- Name: phu_luc_gia_dinh phu_luc_gia_dinh_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.phu_luc_gia_dinh
    ADD CONSTRAINT phu_luc_gia_dinh_pkey PRIMARY KEY (id);


--
-- Name: quyet_dinh_nhan_su quyet_dinh_nhan_su_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.quyet_dinh_nhan_su
    ADD CONSTRAINT quyet_dinh_nhan_su_pkey PRIMARY KEY (id);


--
-- Name: tp_bsc_strategy_map tp_bsc_strategy_map_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_bsc_strategy_map
    ADD CONSTRAINT tp_bsc_strategy_map_pkey PRIMARY KEY (id);


--
-- Name: tp_employee_dependents tp_employee_dependents_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_dependents
    ADD CONSTRAINT tp_employee_dependents_pkey PRIMARY KEY (id);


--
-- Name: tp_employee_p1_assignment tp_employee_p1_assignment_nhan_vien_id_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p1_assignment
    ADD CONSTRAINT tp_employee_p1_assignment_nhan_vien_id_key UNIQUE (nhan_vien_id);


--
-- Name: tp_employee_p1_assignment tp_employee_p1_assignment_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p1_assignment
    ADD CONSTRAINT tp_employee_p1_assignment_pkey PRIMARY KEY (id);


--
-- Name: tp_employee_p2_score tp_employee_p2_score_nhan_vien_id_thang_nam_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p2_score
    ADD CONSTRAINT tp_employee_p2_score_nhan_vien_id_thang_nam_key UNIQUE (nhan_vien_id, thang, nam);


--
-- Name: tp_employee_p2_score tp_employee_p2_score_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p2_score
    ADD CONSTRAINT tp_employee_p2_score_pkey PRIMARY KEY (id);


--
-- Name: tp_employee_p3_score tp_employee_p3_score_nhan_vien_id_thang_nam_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p3_score
    ADD CONSTRAINT tp_employee_p3_score_nhan_vien_id_thang_nam_key UNIQUE (nhan_vien_id, thang, nam);


--
-- Name: tp_employee_p3_score tp_employee_p3_score_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_employee_p3_score
    ADD CONSTRAINT tp_employee_p3_score_pkey PRIMARY KEY (id);


--
-- Name: tp_functional_matrix tp_functional_matrix_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_functional_matrix
    ADD CONSTRAINT tp_functional_matrix_pkey PRIMARY KEY (id);


--
-- Name: tp_job_description tp_job_description_ma_chuc_danh_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_description
    ADD CONSTRAINT tp_job_description_ma_chuc_danh_key UNIQUE (ma_chuc_danh);


--
-- Name: tp_job_description tp_job_description_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_description
    ADD CONSTRAINT tp_job_description_pkey PRIMARY KEY (id);


--
-- Name: tp_job_evaluation tp_job_evaluation_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_evaluation
    ADD CONSTRAINT tp_job_evaluation_pkey PRIMARY KEY (id);


--
-- Name: tp_p1_salary_scale tp_p1_salary_scale_ma_ngach_bac_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p1_salary_scale
    ADD CONSTRAINT tp_p1_salary_scale_ma_ngach_bac_key UNIQUE (ma_ngach, bac);


--
-- Name: tp_p1_salary_scale tp_p1_salary_scale_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p1_salary_scale
    ADD CONSTRAINT tp_p1_salary_scale_pkey PRIMARY KEY (id);


--
-- Name: tp_p2_competency tp_p2_competency_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p2_competency
    ADD CONSTRAINT tp_p2_competency_pkey PRIMARY KEY (id);


--
-- Name: tp_p3_kpi_system tp_p3_kpi_system_ma_kpi_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p3_kpi_system
    ADD CONSTRAINT tp_p3_kpi_system_ma_kpi_key UNIQUE (ma_kpi);


--
-- Name: tp_p3_kpi_system tp_p3_kpi_system_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_p3_kpi_system
    ADD CONSTRAINT tp_p3_kpi_system_pkey PRIMARY KEY (id);


--
-- Name: tp_payroll_config tp_payroll_config_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_config
    ADD CONSTRAINT tp_payroll_config_pkey PRIMARY KEY (id);


--
-- Name: tp_payroll_detail tp_payroll_detail_period_id_nhan_vien_id_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_detail
    ADD CONSTRAINT tp_payroll_detail_period_id_nhan_vien_id_key UNIQUE (period_id, nhan_vien_id);


--
-- Name: tp_payroll_detail tp_payroll_detail_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_detail
    ADD CONSTRAINT tp_payroll_detail_pkey PRIMARY KEY (id);


--
-- Name: tp_payroll_period tp_payroll_period_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_period
    ADD CONSTRAINT tp_payroll_period_pkey PRIMARY KEY (id);


--
-- Name: tp_payroll_period tp_payroll_period_thang_nam_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_period
    ADD CONSTRAINT tp_payroll_period_thang_nam_key UNIQUE (thang, nam);


--
-- Name: tp_policy_version tp_policy_version_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_policy_version
    ADD CONSTRAINT tp_policy_version_pkey PRIMARY KEY (id);


--
-- Name: vi_tri_cong_tac vi_tri_cong_tac_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.vi_tri_cong_tac
    ADD CONSTRAINT vi_tri_cong_tac_pkey PRIMARY KEY (id);


--
-- Name: yeu_cau_reset_mk yeu_cau_reset_mk_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.yeu_cau_reset_mk
    ADD CONSTRAINT yeu_cau_reset_mk_pkey PRIMARY KEY (id);


--
-- Name: idx_cong_van_den_ngay; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_cong_van_den_ngay ON public.cong_van_den USING btree (ngay_den);


--
-- Name: idx_cong_van_den_so; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_cong_van_den_so ON public.cong_van_den USING btree (so_cong_van);


--
-- Name: idx_cong_van_di_ngay; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_cong_van_di_ngay ON public.cong_van_di USING btree (ngay_phat_hanh);


--
-- Name: idx_cong_van_di_so; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_cong_van_di_so ON public.cong_van_di USING btree (so_cong_van);


--
-- Name: idx_danh_muc_phuong_xa_tinh; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_danh_muc_phuong_xa_tinh ON public.danh_muc_phuong_xa USING btree (ma_tinh);


--
-- Name: idx_ho_so_nhan_vien_nv; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_ho_so_nhan_vien_nv ON public.ho_so_nhan_vien USING btree (nhan_vien_id);


--
-- Name: idx_hop_dong_kinh_te_ngay; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_hop_dong_kinh_te_ngay ON public.hop_dong_kinh_te USING btree (ngay_ky);


--
-- Name: idx_hop_dong_kinh_te_so; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_hop_dong_kinh_te_so ON public.hop_dong_kinh_te USING btree (so_hop_dong);


--
-- Name: idx_lich_su_cong_tac_nv; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lich_su_cong_tac_nv ON public.lich_su_cong_tac USING btree (nhan_vien_id);


--
-- Name: idx_lich_su_gui_loi_chuc_loai; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lich_su_gui_loi_chuc_loai ON public.lich_su_gui_loi_chuc USING btree (loai_chuc);


--
-- Name: idx_lich_su_gui_loi_chuc_ngay_gui; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lich_su_gui_loi_chuc_ngay_gui ON public.lich_su_gui_loi_chuc USING btree (ngay_gui);


--
-- Name: idx_lich_su_gui_loi_chuc_nv_id; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lich_su_gui_loi_chuc_nv_id ON public.lich_su_gui_loi_chuc USING btree (nhan_vien_id);


--
-- Name: idx_lsct_nhan_vien_id; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lsct_nhan_vien_id ON public.lich_su_cong_tac USING btree (nhan_vien_id);


--
-- Name: idx_lsct_tu_ngay; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_lsct_tu_ngay ON public.lich_su_cong_tac USING btree (tu_ngay);


--
-- Name: idx_nhan_vien_ten_dang_nhap; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX idx_nhan_vien_ten_dang_nhap ON public.nhan_vien USING btree (ten_dang_nhap) WHERE (ten_dang_nhap IS NOT NULL);


--
-- Name: idx_nhan_vien_trang_thai; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_nhan_vien_trang_thai ON public.nhan_vien USING btree (trang_thai);


--
-- Name: idx_qdns_ngay_quyet_dinh; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_qdns_ngay_quyet_dinh ON public.quyet_dinh_nhan_su USING btree (ngay_quyet_dinh);


--
-- Name: idx_qdns_nhan_vien_id; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_qdns_nhan_vien_id ON public.quyet_dinh_nhan_su USING btree (nhan_vien_id);


--
-- Name: idx_quyet_dinh_nhan_su_nv; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_quyet_dinh_nhan_su_nv ON public.quyet_dinh_nhan_su USING btree (nhan_vien_id);


--
-- Name: idx_ung_vien_trang_thai; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX idx_ung_vien_trang_thai ON public.ung_vien USING btree (trang_thai);


--
-- Name: cham_cong cham_cong_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.cham_cong
    ADD CONSTRAINT cham_cong_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: chat_messages chat_messages_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_messages
    ADD CONSTRAINT chat_messages_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE SET NULL;


--
-- Name: chat_messages chat_messages_room_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_messages
    ADD CONSTRAINT chat_messages_room_id_fkey FOREIGN KEY (room_id) REFERENCES public.chat_rooms(id) ON DELETE CASCADE;


--
-- Name: chat_participants chat_participants_room_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_participants
    ADD CONSTRAINT chat_participants_room_id_fkey FOREIGN KEY (room_id) REFERENCES public.chat_rooms(id) ON DELETE CASCADE;


--
-- Name: chat_participants chat_participants_user_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.chat_participants
    ADD CONSTRAINT chat_participants_user_id_fkey FOREIGN KEY (user_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: danh_muc_phuong_xa danh_muc_phuong_xa_ma_tinh_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.danh_muc_phuong_xa
    ADD CONSTRAINT danh_muc_phuong_xa_ma_tinh_fkey FOREIGN KEY (ma_tinh) REFERENCES public.danh_muc_tinh(ma_tinh);


--
-- Name: ho_so_nhan_vien ho_so_nhan_vien_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.ho_so_nhan_vien
    ADD CONSTRAINT ho_so_nhan_vien_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: lich_su_cong_tac lich_su_cong_tac_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_cong_tac
    ADD CONSTRAINT lich_su_cong_tac_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: lich_su_gui_loi_chuc lich_su_gui_loi_chuc_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.lich_su_gui_loi_chuc
    ADD CONSTRAINT lich_su_gui_loi_chuc_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: phu_luc_gia_dinh phu_luc_gia_dinh_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.phu_luc_gia_dinh
    ADD CONSTRAINT phu_luc_gia_dinh_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: quyet_dinh_nhan_su quyet_dinh_nhan_su_nhan_vien_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.quyet_dinh_nhan_su
    ADD CONSTRAINT quyet_dinh_nhan_su_nhan_vien_id_fkey FOREIGN KEY (nhan_vien_id) REFERENCES public.nhan_vien(id) ON DELETE CASCADE;


--
-- Name: tp_functional_matrix tp_functional_matrix_bsc_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_functional_matrix
    ADD CONSTRAINT tp_functional_matrix_bsc_id_fkey FOREIGN KEY (bsc_id) REFERENCES public.tp_bsc_strategy_map(id) ON DELETE SET NULL;


--
-- Name: tp_job_evaluation tp_job_evaluation_ma_chuc_danh_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_job_evaluation
    ADD CONSTRAINT tp_job_evaluation_ma_chuc_danh_fkey FOREIGN KEY (ma_chuc_danh) REFERENCES public.tp_job_description(ma_chuc_danh) ON DELETE CASCADE;


--
-- Name: tp_payroll_detail tp_payroll_detail_period_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tp_payroll_detail
    ADD CONSTRAINT tp_payroll_detail_period_id_fkey FOREIGN KEY (period_id) REFERENCES public.tp_payroll_period(id) ON DELETE CASCADE;


--
-- Name: cau_hinh_cong_van; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.cau_hinh_cong_van ENABLE ROW LEVEL SECURITY;

--
-- Name: cau_hinh_he_thong; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.cau_hinh_he_thong ENABLE ROW LEVEL SECURITY;

--
-- Name: cham_cong; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.cham_cong ENABLE ROW LEVEL SECURITY;

--
-- Name: chat_messages; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.chat_messages ENABLE ROW LEVEL SECURITY;

--
-- Name: chat_participants; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.chat_participants ENABLE ROW LEVEL SECURITY;

--
-- Name: chat_rooms; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.chat_rooms ENABLE ROW LEVEL SECURITY;

--
-- Name: chuc_danh_ung_vien; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.chuc_danh_ung_vien ENABLE ROW LEVEL SECURITY;

--
-- Name: chuc_vu_danh_muc; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.chuc_vu_danh_muc ENABLE ROW LEVEL SECURITY;

--
-- Name: cong_van_den; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.cong_van_den ENABLE ROW LEVEL SECURITY;

--
-- Name: cong_van_di; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.cong_van_di ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_loai_cong_van; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_loai_cong_van ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_loai_hop_dong; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_loai_hop_dong ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_phong_ban; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_phong_ban ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_phuong_xa; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_phuong_xa ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_tinh; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_tinh ENABLE ROW LEVEL SECURITY;

--
-- Name: danh_muc_trinh_do_hoc_van; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.danh_muc_trinh_do_hoc_van ENABLE ROW LEVEL SECURITY;

--
-- Name: ho_so_nhan_vien; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.ho_so_nhan_vien ENABLE ROW LEVEL SECURITY;

--
-- Name: hop_dong_kinh_te; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.hop_dong_kinh_te ENABLE ROW LEVEL SECURITY;

--
-- Name: lich_su_cong_tac; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.lich_su_cong_tac ENABLE ROW LEVEL SECURITY;

--
-- Name: lich_su_gui_loi_chuc; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.lich_su_gui_loi_chuc ENABLE ROW LEVEL SECURITY;

--
-- Name: mau_dieu_hop_dong; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.mau_dieu_hop_dong ENABLE ROW LEVEL SECURITY;

--
-- Name: nhan_vien; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.nhan_vien ENABLE ROW LEVEL SECURITY;

--
-- Name: phu_luc_gia_dinh; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.phu_luc_gia_dinh ENABLE ROW LEVEL SECURITY;

--
-- Name: quyet_dinh_nhan_su; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.quyet_dinh_nhan_su ENABLE ROW LEVEL SECURITY;

--
-- Name: ung_vien; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.ung_vien ENABLE ROW LEVEL SECURITY;

--
-- Name: vi_tri_cong_tac; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.vi_tri_cong_tac ENABLE ROW LEVEL SECURITY;

--
-- Name: yeu_cau_reset_mk; Type: ROW SECURITY; Schema: public; Owner: -
--

ALTER TABLE public.yeu_cau_reset_mk ENABLE ROW LEVEL SECURITY;

--
-- PostgreSQL database dump complete
--


