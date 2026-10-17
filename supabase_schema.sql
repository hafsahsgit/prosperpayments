-- =========================================================================
-- PROSPER MUSIC ACADEMY / PROSPERWORK
-- Supabase Realtime PostgreSQL Database Schema
-- Run this script in the Supabase SQL Editor (https://supabase.com/dashboard)
-- =========================================================================

-- 1. Enable UUID Extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 2. Students Table
CREATE TABLE IF NOT EXISTS public.students (
    id TEXT PRIMARY KEY DEFAULT ('SCH-' || FLOOR(1000 + RANDOM() * 9000)::TEXT),
    child_name TEXT NOT NULL,
    parent_name TEXT,
    whatsapp TEXT,
    teacher TEXT,
    days TEXT,
    time TEXT,
    status TEXT DEFAULT 'ACTIVE',
    lesson_start TEXT,
    admin TEXT DEFAULT 'Not Yet Assigned',
    meet_link TEXT DEFAULT 'https://meet.google.com/hds-oijm-vkn',
    zoom_link TEXT,
    rate NUMERIC DEFAULT 15,
    fee NUMERIC DEFAULT 60,
    attended INT DEFAULT 0,
    missed INT DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Parents Table
CREATE TABLE IF NOT EXISTS public.parents (
    id TEXT PRIMARY KEY DEFAULT ('PAR-' || FLOOR(1000 + RANDOM() * 9000)::TEXT),
    name TEXT NOT NULL,
    relationship TEXT DEFAULT 'Mother',
    phone TEXT,
    email TEXT,
    address TEXT,
    student_ids JSONB DEFAULT '[]'::JSONB,
    status TEXT DEFAULT 'ACTIVE',
    notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 4. Rescheduled Classes & Makeup Lessons Table
CREATE TABLE IF NOT EXISTS public.reschedules (
    id TEXT PRIMARY KEY DEFAULT ('RES-' || FLOOR(1000 + RANDOM() * 9000)::TEXT),
    student_id TEXT,
    student_name TEXT NOT NULL,
    parent_name TEXT,
    whatsapp TEXT,
    teacher TEXT,
    days TEXT,
    time TEXT,
    original_slot TEXT,
    rescheduled_date DATE,
    admin TEXT DEFAULT 'Not Yet Assigned',
    status TEXT DEFAULT 'RESCHEDULED',
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. Attendance Logs Table
CREATE TABLE IF NOT EXISTS public.attendance_logs (
    id TEXT PRIMARY KEY DEFAULT ('LOG-' || UUID_GENERATE_V4()::TEXT),
    student_id TEXT REFERENCES public.students(id) ON DELETE CASCADE,
    student_name TEXT,
    date DATE NOT NULL,
    status TEXT NOT NULL, -- 'PRESENT' or 'ABSENT'
    marked_by TEXT DEFAULT 'Admin',
    notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 6. Payment Tracker Transactions Table
CREATE TABLE IF NOT EXISTS public.payment_transactions (
    id TEXT PRIMARY KEY DEFAULT ('TXN-' || FLOOR(100000 + RANDOM() * 900000)::TEXT),
    parent_id TEXT,
    parent_name TEXT NOT NULL,
    student_name TEXT,
    amount NUMERIC NOT NULL DEFAULT 0,
    due_date DATE,
    payment_date DATE,
    status TEXT DEFAULT 'UNPAID', -- 'PAID', 'UNPAID', 'PARKED'
    method TEXT DEFAULT 'Bank Transfer',
    notes TEXT,
    parked_month TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 7. System Audit Logs Table
CREATE TABLE IF NOT EXISTS public.audit_logs (
    id UUID PRIMARY KEY DEFAULT UUID_GENERATE_V4(),
    entity TEXT NOT NULL,
    record_id TEXT,
    action TEXT NOT NULL,
    source TEXT DEFAULT 'Web App',
    details TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- =========================================================================
-- ROW LEVEL SECURITY (RLS) POLICIES
-- Enable anonymous read/write access for application API keys
-- =========================================================================

ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.parents ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.reschedules ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.attendance_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.payment_transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.audit_logs ENABLE ROW LEVEL SECURITY;

-- Create permissive policies for anon/authenticated roles
CREATE POLICY "Allow public read/write on students" ON public.students FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read/write on parents" ON public.parents FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read/write on reschedules" ON public.reschedules FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read/write on attendance_logs" ON public.attendance_logs FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read/write on payment_transactions" ON public.payment_transactions FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read/write on audit_logs" ON public.audit_logs FOR ALL USING (true) WITH CHECK (true);

-- =========================================================================
-- REALTIME SUBSCRIPTIONS
-- Enable Supabase Realtime WebSockets for instant 0.1s multi-device updates
-- =========================================================================

BEGIN;
  DROP PUBLICATION IF EXISTS supabase_realtime;
  CREATE PUBLICATION supabase_realtime FOR TABLE 
    public.students, 
    public.parents, 
    public.reschedules, 
    public.attendance_logs, 
    public.payment_transactions;
COMMIT;
